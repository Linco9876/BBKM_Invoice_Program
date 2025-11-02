import json
from typing import Dict, Iterable, List, Optional

import requests

from Outlook_Email_Sorter import GRAPH_BASE, get_access_token, make_session, paged_get


def _headers(access_token: str) -> Dict[str, str]:
    return {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json",
    }


def _get_mail_folder(
    session: requests.Session,
    access_token: str,
    mailbox: str,
    folder_name: str,
) -> Dict:
    if folder_name.lower() == "inbox":
        url = f"{GRAPH_BASE}/users/{mailbox}/mailFolders/inbox"
        response = session.get(url, headers=_headers(access_token), timeout=60)
        if response.status_code != 200:
            raise RuntimeError(
                f"Failed to load Inbox for {mailbox}: {response.status_code} {response.text[:200]}"
            )
        return response.json()

    safe_name = folder_name.replace("'", "''")
    url = (
        f"{GRAPH_BASE}/users/{mailbox}/mailFolders"
        f"?$filter=displayName eq '{safe_name}'&$top=1"
    )
    response = session.get(url, headers=_headers(access_token), timeout=60)
    if response.status_code != 200:
        raise RuntimeError(
            f"Failed to load folder '{folder_name}' for {mailbox}:"
            f" {response.status_code} {response.text[:200]}"
        )

    folders = response.json().get("value", [])
    if not folders:
        raise RuntimeError(f"Folder '{folder_name}' not found for {mailbox}.")
    return folders[0]


def _get_folder_messages(
    session: requests.Session,
    access_token: str,
    mailbox: str,
    folder_id: str,
) -> List[Dict]:
    url = (
        f"{GRAPH_BASE}/users/{mailbox}/mailFolders/{folder_id}/messages"
        "?$orderby=receivedDateTime desc&$select=id,subject,categories"
    )
    return paged_get(session, url, {"Authorization": f"Bearer {access_token}"})


def _matches_category(message_categories: Optional[Iterable[str]], targets: Iterable[str]) -> bool:
    if not message_categories:
        return False
    normalized = {c.strip().lower() for c in message_categories if c}
    targets_normalized = {t.strip().lower() for t in targets if t}
    return bool(normalized & targets_normalized)


def forward_emails_with_categories(to_address: str, categories: Iterable[str]) -> None:
    try:
        session = make_session()
        access_token = get_access_token()

        source_mailbox = "accounts@bbkm.com.au"
        destination_mailbox = to_address

        accounts_inbox = _get_mail_folder(session, access_token, source_mailbox, "Inbox")
        complete_invoices = _get_mail_folder(
            session, access_token, source_mailbox, "Complete Invoices"
        )
        info_inbox = _get_mail_folder(session, access_token, destination_mailbox, "Inbox")

        messages = _get_folder_messages(session, access_token, source_mailbox, accounts_inbox["id"])

        for message in messages:
            subject = message.get("subject") or ""
            categories_current = message.get("categories") or []

            if not _matches_category(categories_current, categories):
                continue

            message_id = message["id"]

            patch_url = f"{GRAPH_BASE}/users/{source_mailbox}/messages/{message_id}"
            patch_resp = session.patch(
                patch_url,
                headers=_headers(access_token),
                data=json.dumps({"isRead": True}),
                timeout=60,
            )
            if patch_resp.status_code not in (200, 202):
                raise RuntimeError(
                    f"Failed to mark message as read: {patch_resp.status_code} {patch_resp.text[:200]}"
                )

            copy_url = f"{GRAPH_BASE}/users/{source_mailbox}/messages/{message_id}/copy"
            copy_body = json.dumps({"destinationId": info_inbox["id"]})
            copy_resp = session.post(
                copy_url,
                headers=_headers(access_token),
                data=copy_body,
                timeout=60,
            )
            if copy_resp.status_code not in (200, 201, 202):
                raise RuntimeError(
                    f"Failed to copy message '{subject}' to {destination_mailbox}:"
                    f" {copy_resp.status_code} {copy_resp.text[:200]}"
                )
            print(f"Copied: '{subject}' to '{info_inbox.get('displayName', 'Inbox')}'.")

            move_url = f"{GRAPH_BASE}/users/{source_mailbox}/messages/{message_id}/move"
            move_body = json.dumps({"destinationId": complete_invoices["id"]})
            move_resp = session.post(
                move_url,
                headers=_headers(access_token),
                data=move_body,
                timeout=60,
            )
            if move_resp.status_code not in (200, 201, 202):
                raise RuntimeError(
                    f"Failed to move message '{subject}' to {complete_invoices.get('displayName')}:"
                    f" {move_resp.status_code} {move_resp.text[:200]}"
                )
            print(f"Moved: '{subject}' to '{complete_invoices.get('displayName', 'Complete Invoices')}'.")

    except Exception as e:
        print(f"Error occurred: {e}")


if __name__ == "__main__":
    forward_emails_with_categories(
        to_address="info@bbkm.com.au",
        categories=[
            "Service Agreement",
            "Reminder",
            "Quote",
            "Remittance",
            "Statement",
            "Caution Email",
            "Credit Adj",
        ],
    )
