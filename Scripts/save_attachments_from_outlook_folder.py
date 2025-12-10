import argparse
import base64
import json
import os
import re
import shutil
import hashlib
from typing import Dict, List, Optional, Tuple

import requests
from msal import ConfidentialClientApplication
from requests.adapters import HTTPAdapter, Retry

DEFAULT_SAVE_PATH = "C:\\BBKM_InvoiceSorter\\Invoices"
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
MANIFEST_PATH = os.path.join(SCRIPT_DIR, "invoice_hashes.json")

GRAPH_SCOPE = ["https://graph.microsoft.com/.default"]
GRAPH_BASE = "https://graph.microsoft.com/v1.0"

USER_EMAIL = os.getenv("OUTLOOK_USER_EMAIL", "accounts@bbkm.com.au")
COMPLETE_FOLDER_NAME = "Completed Invoices"

# Optional hardcoded Azure AD credentials. Populate these fields to force the
# script to use the specified values instead of environment variables or the
# sibling .env file. Leave them blank to keep the existing behavior.
HARD_CODED_AZURE_CREDENTIALS = {
    "AZURE_TENANT_ID": "",
    "AZURE_CLIENT_ID": "",
    "AZURE_CLIENT_SECRET": "",
}

DEFAULT_FORWARD_CATEGORIES = [
    "Service Agreement",
    "Reminder",
    "Quote",
    "Remittance",
    "Statement",
    "Caution Email",
    "Credit Adj",
]


class ForwardingError(RuntimeError):
    """Exception raised when forwarding a message fails."""

    def __init__(self, message: str, *, status_code: Optional[int] = None, permission_denied: bool = False):
        super().__init__(message)
        self.status_code = status_code
        self.permission_denied = permission_denied


class AuthConfigurationError(RuntimeError):
    """Exception raised when authentication cannot succeed without operator action."""


_env_loaded = False
_credential_source = "environment variables"
_credential_usage_logged = False


def _apply_hardcoded_credentials() -> None:
    """Force Azure credentials from in-repo constants when provided."""

    global _credential_source

    for key, value in HARD_CODED_AZURE_CREDENTIALS.items():
        if value:
            os.environ[key] = value
            _credential_source = "HARD_CODED_AZURE_CREDENTIALS"


def _load_env_from_file() -> None:
    """Optionally load Azure credentials from a .env-style file.

    Precedence is as follows:
    1. Hardcoded credentials in ``HARD_CODED_AZURE_CREDENTIALS`` (if set).
    2. If ``AZURE_ENV_FILE`` is set, load from that path (erroring if missing)
       and allow it to override existing environment values.
    3. Otherwise, if a ``.env`` file sits alongside this script, load from it
       and override existing values. This allows operators to drop a sibling
       ``.env`` (for example on Windows deployments) and have it take effect
       even if stale variables are already present in the environment.
    4. Explicitly exported environment variables still take priority if they
       are set *after* this loader runs.

    This keeps secrets out of the repository while still allowing operators to
    provide them locally—especially in environments where a sibling ``.env`` is
    distributed with the scripts.
    """

    global _env_loaded

    if _env_loaded:
        return

    env_path = os.getenv("AZURE_ENV_FILE")
    override_existing = True
    if env_path:
        if not os.path.exists(env_path):
            raise RuntimeError(f"AZURE_ENV_FILE is set but {env_path} does not exist")
        source_label = f"AZURE_ENV_FILE ({env_path})"
    else:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        env_path = os.path.join(script_dir, ".env")
        if not os.path.exists(env_path):
            return
        source_label = f".env next to scripts ({env_path})"

    with open(env_path, "r", encoding="utf-8") as env_file:
        for raw_line in env_file:
            line = raw_line.strip()
            if not line or line.startswith("#"):
                continue
            if "=" not in line:
                continue
            key, value = line.split("=", 1)
            key = key.strip()
            value = value.strip()
            if override_existing or key not in os.environ:
                os.environ[key] = value

    _env_loaded = True
    global _credential_source
    _credential_source = source_label


def _log_credential_usage(tenant_id: str, client_id: str) -> None:
    """Emit a one-time log entry showing which credentials are active."""

    global _credential_usage_logged
    if _credential_usage_logged:
        return

    source = _credential_source or "environment variables"
    tenant_hint = tenant_id or "(unset)"
    client_hint = client_id or "(unset)"
    print(
        f"Using Azure credentials from {source}: tenant {tenant_hint}, client {client_hint}"
    )
    _credential_usage_logged = True

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
        if response.status_code in (200, 202, 404):
            # 404 is treated as a benign condition: the message may have been moved
            # or deleted between retrieval and update attempts.
            self._dirty = False
            return
        if response.status_code not in (200, 202):
            raise RuntimeError(
                f"Failed to update message {self._message_id}: {response.status_code} {response.text[:200]}"
            )
        self._dirty = False


def _set_user_email(user_email: str) -> None:
    """Update the target mailbox for subsequent Graph calls."""

    global USER_EMAIL
    USER_EMAIL = (user_email or "").strip() or USER_EMAIL


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
    _load_env_from_file()
    _apply_hardcoded_credentials()
    tenant_id = _ensure_env("AZURE_TENANT_ID")
    client_id = _ensure_env("AZURE_CLIENT_ID")
    authority = f"https://login.microsoftonline.com/{tenant_id}"
    _log_credential_usage(tenant_id, client_id)
    app = ConfidentialClientApplication(
        client_id,
        authority=authority,
        client_credential=_ensure_env("AZURE_CLIENT_SECRET"),
    )
    token = app.acquire_token_for_client(GRAPH_SCOPE)
    if "access_token" not in token:
        error = (token.get("error") or "").lower()
        description = token.get("error_description") or ""
        codes = token.get("error_codes") or []

        if error == "invalid_client" or 7000222 in codes:
            raise AuthConfigurationError(
                "Azure AD client secret appears to be expired or invalid. "
                "Create a new client secret for the app registration and set "
                "AZURE_CLIENT_SECRET before retrying. Details: "
                f"{description or token}"
            )

        raise RuntimeError(f"Failed to acquire token: {token}")
    return token["access_token"]


def _describe_graph_error(
    response: requests.Response, action: str, *, mailbox_hint: Optional[str] = None
) -> str:
    """Return a helpful error string for Microsoft Graph failures."""

    detail = ""
    try:
        payload = response.json()
        error = payload.get("error") or {}
        code = error.get("code")
        message = error.get("message")
        if code or message:
            detail = f"{code or 'Error'}: {message or ''}".strip()
    except Exception:
        # fall back to raw body below
        pass

    if response.status_code in (401, 403):
        mailbox = (mailbox_hint or USER_EMAIL).strip() or USER_EMAIL
        credential_hint = (
            f" Currently using client id {os.getenv('AZURE_CLIENT_ID', '(unset)')}"
            f" from {_credential_source} in tenant {os.getenv('AZURE_TENANT_ID', '(unset)')}"
        )
        hint = (
            " Confirm that AZURE_TENANT_ID, AZURE_CLIENT_ID, and AZURE_CLIENT_SECRET"
            " are set for the app registration and that it has Mail.ReadWrite and"
            f" Mail.Send application permissions for {mailbox}. If the mailbox"
            " differs, pass --user-email or set OUTLOOK_USER_EMAIL to the target"
            " address (accounts@bbkm.com.au by default). For cross-mailbox copies,"
            " ensure the app registration has rights to the destination mailbox"
            " (e.g., info@bbkm.com.au)."
            f"{credential_hint}."
        )
        detail = f"{detail or response.text[:200]}{hint}"
    else:
        detail = detail or response.text[:200]

    return f"{action}: {response.status_code} {detail}"


def _get_message_details(session: requests.Session, token: str, message_id: str) -> Dict[str, object]:
    url = f"{GRAPH_BASE}/users/{USER_EMAIL}/messages/{message_id}?$select=id,subject,body,categories,flag,isRead,from"
    headers = {"Authorization": f"Bearer {token}"}
    response = session.get(url, headers=headers, timeout=60)
    if response.status_code != 200:
        raise RuntimeError(_describe_graph_error(response, f"Failed to fetch message {message_id}"))
    return response.json()


def _list_inbox_messages(
    session: requests.Session, token: str, *, newest_first: bool = False
) -> List[Dict[str, object]]:
    order = "desc" if newest_first else "asc"
    url = (
        f"{GRAPH_BASE}/users/{USER_EMAIL}/mailFolders/inbox/messages"
        f"?$orderby=receivedDateTime {order}"
        "&$select=id,subject,from,categories,flag,hasAttachments,isRead"
        "&$top=50"
    )
    headers = {"Authorization": f"Bearer {token}"}
    messages: List[Dict[str, object]] = []
    while url:
        response = session.get(url, headers=headers, timeout=60)
        if response.status_code in (401, 403):
            # Access denied will not succeed with retries; surface a fatal auth
            # configuration error so the main loop can halt and the operator can
            # correct mailbox permissions or credentials before retrying.
            raise AuthConfigurationError(
                _describe_graph_error(response, "Failed to list messages")
            )
        if response.status_code != 200:
            raise RuntimeError(_describe_graph_error(response, "Failed to list messages"))
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
            raise RuntimeError(_describe_graph_error(response, "Failed to list attachments"))
        payload = response.json()
        attachments.extend(payload.get("value", []))
        url = payload.get("@odata.nextLink")
    return attachments


def _find_mail_folder_id(
    session: requests.Session,
    token: str,
    *,
    display_name: str,
    user_email: Optional[str] = None,
) -> str:
    """Return the folder id that matches ``display_name`` (case-insensitive)."""

    mailbox = (user_email or USER_EMAIL).strip()
    target = display_name.casefold()
    url = f"{GRAPH_BASE}/users/{mailbox}/mailFolders?$top=200"
    headers = {"Authorization": f"Bearer {token}"}

    while url:
        response = session.get(url, headers=headers, timeout=60)
        if response.status_code != 200:
            raise RuntimeError(
                _describe_graph_error(
                    response,
                    "Failed to list mail folders",
                    mailbox_hint=mailbox,
                )
            )

        payload = response.json()
        for folder in payload.get("value", []):
            name = (folder.get("displayName") or "").casefold()
            if name == target:
                folder_id = folder.get("id")
                if folder_id:
                    return folder_id

        url = payload.get("@odata.nextLink")
    raise RuntimeError(f"Unable to locate folder named '{display_name}' for {mailbox}")


def _move_message(
    session: requests.Session,
    token: str,
    message_id: str,
    destination_id: str,
    *,
    user_email: Optional[str] = None,
) -> None:
    mailbox = (user_email or USER_EMAIL).strip()
    url = f"{GRAPH_BASE}/users/{mailbox}/messages/{message_id}/move"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }
    body = json.dumps({"destinationId": destination_id})
    response = session.post(url, headers=headers, data=body, timeout=60)
    if response.status_code in (200, 201):
        return

    if response.status_code == 404:
        # Message may have been moved or deleted by another agent between listing and action.
        return

    raise RuntimeError(
        _describe_graph_error(
            response,
            f"Failed to move message {message_id}",
            mailbox_hint=mailbox,
        )
    )


def _delete_message(
    session: requests.Session, token: str, message_id: str, *, user_email: Optional[str] = None
) -> None:
    mailbox = (user_email or USER_EMAIL).strip()
    url = f"{GRAPH_BASE}/users/{mailbox}/messages/{message_id}"
    headers = {"Authorization": f"Bearer {token}"}
    response = session.delete(url, headers=headers, timeout=60)
    if response.status_code in (200, 204):
        return
    if response.status_code == 404:
        return
    raise RuntimeError(
        _describe_graph_error(
            response,
            f"Failed to delete message {message_id}",
            mailbox_hint=mailbox,
        )
    )


def _ensure_attachment_content(
    session: requests.Session,
    token: str,
    message_id: str,
    attachment: Dict[str, object],
) -> str:
    content = attachment.get("contentBytes")
    if content:
        return str(content)

    attachment_id = attachment.get("id")
    if not attachment_id:
        raise RuntimeError("Attachment is missing both inline content and an identifier")

    att_url = (
        f"{GRAPH_BASE}/users/{USER_EMAIL}/messages/{message_id}/attachments/{attachment_id}/$value"
    )
    headers = {"Authorization": f"Bearer {token}"}
    response = session.get(att_url, headers=headers, timeout=60)
    if response.status_code != 200:
        raise RuntimeError(
            _describe_graph_error(
                response, f"Failed to download attachment {attachment_id}"
            )
        )
    return base64.b64encode(response.content).decode("ascii")


def _copy_message_to_mailbox(
    session: requests.Session,
    token: str,
    email: GraphEmailProxy,
    attachments: List[Dict[str, object]],
    target_mailbox: str,
    target_folder: str = "Inbox",
) -> None:
    mailbox = (target_mailbox or "").strip()
    if not mailbox:
        return

    folder_id = _find_mail_folder_id(
        session, token, display_name=target_folder, user_email=mailbox
    )

    message_attachments: List[Dict[str, object]] = []
    for attachment in attachments:
        odata_type = (attachment.get("@odata.type") or "").lower()
        if "fileattachment" not in odata_type:
            continue
        name = attachment.get("name") or "attachment"
        content_type = attachment.get("contentType") or "application/octet-stream"
        content_bytes = _ensure_attachment_content(session, token, email.id, attachment)
        message_attachments.append(
            {
                "@odata.type": "#microsoft.graph.fileAttachment",
                "name": name,
                "contentType": content_type,
                "contentBytes": content_bytes,
            }
        )

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }
    body: Dict[str, object] = {
        "subject": email.Subject,
        "body": {"contentType": "HTML", "content": email.Body},
        "toRecipients": [{"emailAddress": {"address": mailbox}}],
        "isRead": False,
    }
    if message_attachments:
        body["attachments"] = message_attachments

    url = f"{GRAPH_BASE}/users/{mailbox}/mailFolders/{folder_id}/messages"
    response = session.post(url, headers=headers, data=json.dumps(body), timeout=60)
    if response.status_code in (200, 201):
        return

    raise RuntimeError(
        _describe_graph_error(
            response,
            f"Failed to copy message '{email.Subject}' to {mailbox} {target_folder}",
            mailbox_hint=mailbox,
        )
    )


def _send_message_copy(
    session: requests.Session,
    token: str,
    email: GraphEmailProxy,
    to_address: str,
    attachments: List[Dict[str, object]],
) -> None:
    mail_attachments: List[Dict[str, object]] = []

    for attachment in attachments:
        odata_type = (attachment.get("@odata.type") or "").lower()
        if "fileattachment" not in odata_type:
            continue

        name = attachment.get("name") or "attachment"
        content_type = attachment.get("contentType") or "application/octet-stream"
        content_bytes = _ensure_attachment_content(session, token, email.id, attachment)

        mail_attachments.append(
            {
                "@odata.type": "#microsoft.graph.fileAttachment",
                "name": name,
                "contentType": content_type,
                "contentBytes": content_bytes,
            }
        )

    send_url = f"{GRAPH_BASE}/users/{USER_EMAIL}/sendMail"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }
    body = {
        "message": {
            "subject": email.Subject,
            "body": {"contentType": "HTML", "content": email.Body},
            "toRecipients": [{"emailAddress": {"address": to_address}}],
        },
        "saveToSentItems": False,
    }

    if mail_attachments:
        body["message"]["attachments"] = mail_attachments

    response = session.post(send_url, headers=headers, data=json.dumps(body), timeout=60)
    if response.status_code in (202, 200):
        return

    if response.status_code == 403:
        forward_url = f"{GRAPH_BASE}/users/{USER_EMAIL}/messages/{email.id}/forward"
        forward_headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
        forward_body = {
            "comment": "",
            "toRecipients": [{"emailAddress": {"address": to_address}}],
        }
        forward_response = session.post(
            forward_url, headers=forward_headers, data=json.dumps(forward_body), timeout=60
        )
        if forward_response.status_code in (202, 200, 204):
            return

        raise ForwardingError(
            _describe_graph_error(
                response,
                f"Failed to send message copy to {to_address}",
                mailbox_hint=USER_EMAIL,
            )
            + " and forward fallback "
            + _describe_graph_error(
                forward_response,
                f"Failed to forward message directly to {to_address}",
                mailbox_hint=USER_EMAIL,
            ),
            status_code=403,
            permission_denied=True,
        )

    raise ForwardingError(
        _describe_graph_error(
            response,
            f"Failed to send message copy to {to_address}",
            mailbox_hint=USER_EMAIL,
        ),
        status_code=response.status_code,
    )


def _strip_html(text: str) -> str:
    return re.sub(r"<[^>]+>", " ", text or "")


def _fetch_attachment_bytes(
    session: requests.Session,
    token: str,
    email_id: str,
    attachment: Dict[str, object],
) -> Optional[bytes]:
    """Return raw bytes for an attachment, downloading if needed."""

    content = attachment.get("contentBytes")
    if content:
        try:
            return base64.b64decode(content)
        except Exception:
            return None

    attachment_id = attachment.get("id")
    if not attachment_id:
        return None

    att_url = f"{GRAPH_BASE}/users/{USER_EMAIL}/messages/{email_id}/attachments/{attachment_id}/$value"
    headers = {"Authorization": f"Bearer {token}"}
    response = session.get(att_url, headers=headers, timeout=60)
    if response.status_code != 200:
        print(f"Error downloading attachment stream: {response.status_code}")
        return None

    return response.content


def _compute_hash(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _compute_file_hash(path: str) -> Optional[str]:
    if not os.path.exists(path):
        return None
    with open(path, "rb") as handle:
        return _compute_hash(handle.read())


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

    manifest_path = MANIFEST_PATH
    seen_hashes: Dict[str, str] = {}
    if os.path.exists(manifest_path):
        try:
            with open(manifest_path, "r", encoding="utf-8") as handle:
                seen_hashes = json.load(handle) or {}
        except Exception:
            seen_hashes = {}
    manifest_dirty = False

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
        "Caution Email",
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
            continue

        if re.search(r"\breminder\b", subject, re.IGNORECASE) or re.search(r"\breminder\b", body_text, re.IGNORECASE):
            email.Categories = "Reminder"
            email.Save()
            continue

        if re.search(r"\bquote\b", subject, re.IGNORECASE) or re.search(r"\bquote\b", body_text, re.IGNORECASE):
            email.Categories = "Quote"
            email.Save()
            continue

        if re.search(r"\bOver-Due\b", subject, re.IGNORECASE) or re.search(r"\bOverdue\b", subject, re.IGNORECASE):
            email.Categories = "Reminder"
            email.Save()
            continue

        if re.search(r"\bStatement\b", subject, re.IGNORECASE) and not re.search(r"\bActivity Statement\b", subject, re.IGNORECASE):
            email.Categories = "Statement"
            email.Save()
            continue

        if re.search(r"\bCredit Adj\b", subject, re.IGNORECASE) or re.search(r"\bCredit Adj\b", body_text, re.IGNORECASE):
            email.Categories = "Credit Adj"
            email.Save()
            continue

        if "24 Pritchard Street" in subject or "24 Pritchard Street" in body_text:
            if "Activity Statement" not in subject:
                email.Categories = "Skipped Email"
                email.Save()
                continue

        if any(domain in sender_email for domain in ["bbkm.com.au"]):
            if "Activity Statement" not in subject:
                email.Categories = "Skipped Email"
                email.Save()
                continue

        attachments: List[Dict[str, object]] = []
        if basic_message.get("hasAttachments"):
            attachments = _list_attachments(session, token, email.id)

        skip_email_due_to_remittance = False
        for attachment in attachments:
            if re.search(r"remittance", (attachment.get("name") or ""), re.IGNORECASE):
                email.Categories = "Remittance"
                email.Save()
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

            data = _fetch_attachment_bytes(session, token, email.id, attachment)
            if data is None:
                continue

            content_hash = _compute_hash(data)
            if content_hash in seen_hashes:
                email.Categories = "Doubled up"
                email.Save()
                continue

            business_name = next((name_hint for domain, name_hint in domain_map.items() if domain in sender_email), None)
            base_name, ext = os.path.splitext(name)
            file_name = f"{base_name} {business_name}{ext}" if business_name else name
            destination_path = os.path.join(save_path, file_name)

            if os.path.exists(destination_path):
                existing_hash = _compute_file_hash(destination_path)
                if existing_hash and existing_hash == content_hash:
                    email.Categories = "Doubled up"
                    email.Save()
                    seen_hashes[content_hash] = destination_path
                    manifest_dirty = True
                    continue

                destination_path = os.path.join(save_path, f"new_{name}")

            with open(destination_path, "wb") as handle:
                handle.write(data)

            attachment_saved = True
            saved_attachments.append((email, destination_path))
            email_file_map[name] = email
            seen_hashes[content_hash] = destination_path
            manifest_dirty = True

        if not has_attachments:
            email.FlagStatus = 2
            email.Save()
        elif attachment_saved:
            email.Categories = "Attachment Extracted"
            email.UnRead = False
            email.Save()

        break

    if manifest_dirty:
        try:
            with open(manifest_path, "w", encoding="utf-8") as handle:
                json.dump(seen_hashes, handle, indent=2)
        except Exception:
            pass

    return saved_attachments, email_file_map


def forward_emails_with_categories(
    to_address: str,
    categories: List[str],
    *,
    post_forward_mailbox: Optional[str] = None,
    post_forward_folder: str = "Inbox",
) -> None:
    """Skip categorised emails instead of forwarding or copying them."""

    if not categories:
        return

    token = _get_access_token()
    session = _make_session()

    category_targets = {category.casefold() for category in categories}

    messages = _list_inbox_messages(session, token, newest_first=True)

    for basic_message in messages:
        raw_categories = basic_message.get("categories")
        if not raw_categories:
            continue

        if isinstance(raw_categories, list):
            message_categories = {str(value).casefold() for value in raw_categories}
        else:
            message_categories = {str(raw_categories).casefold()}

        if not (message_categories & category_targets):
            continue

        continue



def main(argv: Optional[List[str]] = None) -> None:
    """Command line entry point for saving attachments or forwarding emails."""

    parser = argparse.ArgumentParser(
        description=(
            "Download attachments from the shared inbox or skip categorised "
            "emails (forwarding disabled)."
        )
    )
    parser.add_argument(
        "--forward",
        action="store_true",
        help=(
            "Skip emails that match the supplied categories instead of "
            "downloading attachments."
        ),
    )
    parser.add_argument(
        "--to-address",
        default="",
        help=(
            "Destination email address for forwarding (currently unused; "
            "categorised emails remain in the source mailbox)."
        ),
    )
    parser.add_argument(
        "--category",
        dest="categories",
        action="append",
        metavar="NAME",
        help=(
            "Category to match when forwarding emails. Provide multiple times "
            "to include more than one category. Defaults to the standard "
            "processing list."
        ),
    )
    parser.add_argument(
        "--folder",
        default="invoices",
        help="Source folder name when downloading attachments.",
    )
    parser.add_argument(
        "--destination",
        default=DEFAULT_SAVE_PATH,
        metavar="PATH",
        help="Filesystem path for saving attachments.",
    )
    parser.add_argument(
        "--user-email",
        default=os.getenv("OUTLOOK_USER_EMAIL", USER_EMAIL),
        help=(
            "Mailbox to process with Microsoft Graph. Overrides OUTLOOK_USER_EMAIL "
            "and defaults to accounts@bbkm.com.au."
        ),
    )
    parser.add_argument(
        "--post-forward-mailbox",
        default="",
        help=(
            "Mailbox to hold forwarded messages (unused while forwarding is "
            "disabled)."
        ),
    )
    parser.add_argument(
        "--post-forward-folder",
        default="Inbox",
        help=(
            "Folder name in the post-forward mailbox to store the moved message. "
            "Defaults to Inbox."
        ),
    )

    args = parser.parse_args(argv)

    _set_user_email(args.user_email)

    if args.forward:
        categories = args.categories or DEFAULT_FORWARD_CATEGORIES
        forward_emails_with_categories(
            args.to_address,
            categories,
            post_forward_mailbox=args.post_forward_mailbox,
            post_forward_folder=args.post_forward_folder,
        )
        return

    save_attachments_from_outlook_folder(args.folder, args.destination)


if __name__ == "__main__":
    main()
