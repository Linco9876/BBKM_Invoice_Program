import base64
import json
import os
import re
import shutil
import filecmp
from typing import Dict, List, Tuple

import requests
from msal import ConfidentialClientApplication
from requests.adapters import HTTPAdapter, Retry

save_path = "C:\\BBKM_InvoiceSorter\\Invoices"

GRAPH_SCOPE = ["https://graph.microsoft.com/.default"]
GRAPH_BASE = "https://graph.microsoft.com/v1.0"

USER_EMAIL = os.getenv("OUTLOOK_USER_EMAIL", "accounts@bbkm.com.au")


class GraphEmailProxy:
    """Lightweight wrapper to mimic the few Outlook MailItem members we rely on."""

    def __init__(self, session: requests.Session, access_token: str, message: Dict[str, object]):
        self._session = session
        self._access_token = access_token
        self._message_id = message.get("id")
        self.Subject = message.get("subject", "") or ""
        body = message.get("body", {}) or {}
        self.Body = body.get("content", "") or ""
        sender = ((message.get("from") or {}).get("emailAddress") or {}).get("address")
        self.SenderEmailAddress = (sender or "").lower()
        categories = message.get("categories") or []
        if isinstance(categories, list):
            self._categories = list(categories)
        elif categories:
            self._categories = [str(categories)]
        else:
            self._categories = []
        flag = message.get("flag") or {}
        self._flag_state = flag.get("flagStatus")
        is_read = bool(message.get("isRead", False))
        self._is_unread = not is_read
        self._dirty = False
        self.IsConflict = False

    @property
    def id(self) -> str:
        return self._message_id  # type: ignore[return-value]

    @property
    def Categories(self) -> str:
        return ", ".join(self._categories)

    @Categories.setter
    def Categories(self, value: str) -> None:
        self._categories = [value] if value else []
        self._dirty = True

    @property
    def FlagStatus(self) -> int:
        return 2 if self._flag_state == "complete" else 0

    @FlagStatus.setter
    def FlagStatus(self, value: int) -> None:
        self._flag_state = "complete" if value == 2 else "notFlagged"
        self._dirty = True

    @property
    def UnRead(self) -> bool:
        return self._is_unread

    @UnRead.setter
    def UnRead(self, value: bool) -> None:
        self._is_unread = bool(value)
        self._dirty = True

    def Save(self) -> None:
        if not self._dirty:
            return
        url = f"{GRAPH_BASE}/users/{USER_EMAIL}/messages/{self._message_id}"
        headers = {
            "Authorization": f"Bearer {self._access_token}",
            "Content-Type": "application/json",
        }
        body: Dict[str, object] = {
            "categories": self._categories,
            "isRead": not self._is_unread,
        }
        if self._flag_state:
            body["flag"] = {"flagStatus": self._flag_state}
        response = self._session.patch(url, headers=headers, data=json.dumps(body), timeout=60)
        if response.status_code not in (200, 202):
            raise RuntimeError(
                f"Failed to update message {self._message_id}: {response.status_code} {response.text[:200]}"
            )
        self._dirty = False


def _ensure_env(var_name: str) -> str:
    value = os.getenv(var_name)
    if not value:
        raise RuntimeError(f"Missing required environment variable: {var_name}")
    return value


def _make_session() -> requests.Session:
    session = requests.Session()
    retries = Retry(
        total=5,
        connect=5,
        read=5,
        backoff_factor=0.5,
        status_forcelist=(429, 500, 502, 503, 504),
        allowed_methods=frozenset(["GET", "POST", "PATCH"]),
    )
    adapter = HTTPAdapter(max_retries=retries, pool_connections=10, pool_maxsize=10)
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    return session


def _get_access_token() -> str:
    authority = f"https://login.microsoftonline.com/{_ensure_env('AZURE_TENANT_ID')}"
    app = ConfidentialClientApplication(
        _ensure_env("AZURE_CLIENT_ID"),
        authority=authority,
        client_credential=_ensure_env("AZURE_CLIENT_SECRET"),
    )
    token = app.acquire_token_for_client(GRAPH_SCOPE)
    if "access_token" not in token:
        raise RuntimeError(f"Failed to acquire token: {token}")
    return token["access_token"]


def compare_files(file1: str, file2: str) -> bool:
    if not os.path.exists(file1) or not os.path.exists(file2):
        return False
    return filecmp.cmp(file1, file2, shallow=False)


def _get_message_details(session: requests.Session, token: str, message_id: str) -> Dict[str, object]:
    url = f"{GRAPH_BASE}/users/{USER_EMAIL}/messages/{message_id}?$select=id,subject,body,categories,flag,isRead,from"
    headers = {"Authorization": f"Bearer {token}"}
    response = session.get(url, headers=headers, timeout=60)
    if response.status_code != 200:
        raise RuntimeError(f"Failed to fetch message {message_id}: {response.status_code} {response.text[:200]}")
    return response.json()


def _list_inbox_messages(session: requests.Session, token: str) -> List[Dict[str, object]]:
    url = (
        f"{GRAPH_BASE}/users/{USER_EMAIL}/mailFolders/inbox/messages"
        "?$orderby=receivedDateTime asc"
        "&$select=id,subject,from,categories,flag,hasAttachments,isRead"
        "&$top=50"
    )
    headers = {"Authorization": f"Bearer {token}"}
    messages: List[Dict[str, object]] = []
    while url:
        response = session.get(url, headers=headers, timeout=60)
        if response.status_code != 200:
            raise RuntimeError(f"Failed to list messages: {response.status_code} {response.text[:200]}")
        payload = response.json()
        messages.extend(payload.get("value", []))
        url = payload.get("@odata.nextLink")
    return messages


def _list_attachments(session: requests.Session, token: str, message_id: str) -> List[Dict[str, object]]:
    url = f"{GRAPH_BASE}/users/{USER_EMAIL}/messages/{message_id}/attachments"
    headers = {"Authorization": f"Bearer {token}"}
    attachments: List[Dict[str, object]] = []
    while url:
        response = session.get(url, headers=headers, timeout=60)
        if response.status_code != 200:
            raise RuntimeError(f"Failed to list attachments: {response.status_code} {response.text[:200]}")
        payload = response.json()
        attachments.extend(payload.get("value", []))
        url = payload.get("@odata.nextLink")
    return attachments


def _strip_html(text: str) -> str:
    return re.sub(r"<[^>]+>", " ", text or "")


def _save_attachment(content_bytes: str, target_path: str) -> None:
    with open(target_path, "wb") as handle:
        handle.write(base64.b64decode(content_bytes))


def _download_attachment_to_path(
    session: requests.Session,
    token: str,
    email_id: str,
    attachment: Dict[str, object],
    target_path: str,
) -> bool:
    content = attachment.get("contentBytes")
    if content:
        _save_attachment(content, target_path)
        return True

    attachment_id = attachment.get("id")
    if not attachment_id:
        return False

    att_url = f"{GRAPH_BASE}/users/{USER_EMAIL}/messages/{email_id}/attachments/{attachment_id}/$value"
    headers = {"Authorization": f"Bearer {token}"}
    response = session.get(att_url, headers=headers, timeout=60)
    if response.status_code == 200:
        with open(target_path, "wb") as handle:
            handle.write(response.content)
        return True

    print(f"Error downloading attachment stream: {response.status_code}")
    return False


def save_attachments_from_outlook_folder(
    folder_name: str, save_path: str
) -> Tuple[List[Tuple[GraphEmailProxy, str]], Dict[str, GraphEmailProxy]]:
    """Download attachments from the shared inbox using Microsoft Graph.

    Returns a tuple with the same structure as the previous COM-based implementation:
    - A list of ``(email_proxy, file_path)`` tuples for each saved attachment.
    - A mapping of original attachment names to their originating ``email_proxy``.
    """
    token = _get_access_token()
    session = _make_session()

    saved_attachments: List[Tuple[GraphEmailProxy, str]] = []
    email_file_map: Dict[str, GraphEmailProxy] = {}

    os.makedirs(save_path, exist_ok=True)

    messages = _list_inbox_messages(session, token)

    skip_categories = {
        "Attachment Extracted",
        "Skipped Email",
        "Reminder",
        "Service Agreement",
        "Quote",
        "Statement",
        "Credit Adj",
        "Remittance",
        "Doubled up",
    }

    domain_map = {
        "independenceaustralia.com": "independence australia",
        "country-care.com.au": "country care",
        "brightsky.com.au": "brightsky",
        "visionaustralia.org": "visionaustralia",
        "gsc.vic.gov.au": "Gannawarra Shire",
        "alifesimplylived.com.au": "A Life Simply Lived",
    }

    for basic_message in messages:
        raw_categories = basic_message.get("categories")
        if isinstance(raw_categories, list):
            categories = set(raw_categories)
        elif isinstance(raw_categories, str):
            categories = {raw_categories}
        else:
            categories = set()
        if categories & skip_categories:
            continue

        flag = (basic_message.get("flag") or {}).get("flagStatus")
        if flag == "complete":
            continue

        message_details = _get_message_details(session, token, basic_message.get("id", ""))
        email = GraphEmailProxy(session, token, message_details)

        subject = email.Subject
        body_text = _strip_html(email.Body)
        sender_email = email.SenderEmailAddress

        # Keyword-based classification
        if any(re.search(r"\bService Agreement\b", t, re.IGNORECASE) for t in [subject, body_text]):
            email.Categories = "Service Agreement"
            email.Save()
            print("Service Agreement Found")
            continue

        if re.search(r"\breminder\b", subject, re.IGNORECASE) or re.search(r"\breminder\b", body_text, re.IGNORECASE):
            email.Categories = "Reminder"
            email.Save()
            print("Reminder Found")
            continue

        if re.search(r"\bquote\b", subject, re.IGNORECASE) or re.search(r"\bquote\b", body_text, re.IGNORECASE):
            email.Categories = "Quote"
            email.Save()
            print("Quote Found")
            continue

        if re.search(r"\bOver-Due\b", subject, re.IGNORECASE) or re.search(r"\bOverdue\b", subject, re.IGNORECASE):
            email.Categories = "Reminder"
            email.Save()
            print("Reminder Found")
            continue

        if re.search(r"\bStatement\b", subject, re.IGNORECASE) and not re.search(r"\bActivity Statement\b", subject, re.IGNORECASE):
            email.Categories = "Statement"
            email.Save()
            print("Statement Found")
            continue

        if re.search(r"\bCredit Adj\b", subject, re.IGNORECASE) or re.search(r"\bCredit Adj\b", body_text, re.IGNORECASE):
            email.Categories = "Credit Adj"
            email.Save()
            print("Credit Adj Found")
            continue

        if "24 Pritchard Street" in subject or "24 Pritchard Street" in body_text:
            if "Activity Statement" not in subject:
                email.Categories = "Skipped Email"
                email.Save()
                print("Skipped Email Found")
                continue

        if any(domain in sender_email for domain in ["bbkm.com.au"]):
            if "Activity Statement" not in subject:
                email.Categories = "Skipped Email"
                email.Save()
                print("Skipped Domain")
                continue

        attachments: List[Dict[str, object]] = []
        if basic_message.get("hasAttachments"):
            attachments = _list_attachments(session, token, email.id)

        skip_email_due_to_remittance = False
        for attachment in attachments:
            if re.search(r"remittance", (attachment.get("name") or ""), re.IGNORECASE):
                email.Categories = "Remittance"
                email.Save()
                print("Remittance Found")
                skip_email_due_to_remittance = True
                break
        if skip_email_due_to_remittance:
            continue

        attachment_saved = False
        has_attachments = False

        for attachment in attachments:
            name = attachment.get("name") or ""
            odata_type = (attachment.get("@odata.type") or "").lower()
            if "fileattachment" not in odata_type:
                continue
            if not name.lower().endswith((".pdf", ".docx", ".doc")):
                continue

            has_attachments = True

            business_name = next((name_hint for domain, name_hint in domain_map.items() if domain in sender_email), None)
            base_name, ext = os.path.splitext(name)
            file_name = f"{base_name} {business_name}{ext}" if business_name else name
            destination_path = os.path.join(save_path, file_name)

            if os.path.exists(destination_path):
                temp_file_path = os.path.join(save_path, f"temp_{name}")
                if not _download_attachment_to_path(session, token, email.id, attachment, temp_file_path):
                    continue
                if compare_files(destination_path, temp_file_path):
                    os.remove(temp_file_path)
                    email.Categories = "Doubled up"
                    email.Save()
                    continue
                new_path = os.path.join(save_path, f"new_{name}")
                shutil.move(temp_file_path, new_path)
                destination_path = new_path
            else:
                if not _download_attachment_to_path(session, token, email.id, attachment, destination_path):
                    continue

            attachment_saved = True
            saved_attachments.append((email, destination_path))
            email_file_map[name] = email

        if not has_attachments:
            email.FlagStatus = 2
            email.Save()
        elif attachment_saved:
            email.Categories = "Attachment Extracted"
            email.UnRead = False
            email.Save()

        break

    return saved_attachments, email_file_map
