import os
import json
import base64
import time
from typing import Dict, List, Optional

import requests
from requests.adapters import HTTPAdapter, Retry

try:
    from dotenv import load_dotenv  # optional, for local dev via .env
    load_dotenv()
except Exception:
    pass

from msal import ConfidentialClientApplication

# =========================
# Configuration (env-based)
# =========================
TENANT_ID = os.getenv("AZURE_TENANT_ID")
CLIENT_ID = os.getenv("AZURE_CLIENT_ID")
CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET")  # DO NOT hardcode
USER_EMAIL = os.getenv("OUTLOOK_USER_EMAIL", "accounts@bbkm.com.au")

GRAPH_SCOPE = ["https://graph.microsoft.com/.default"]
GRAPH_BASE = "https://graph.microsoft.com/v1.0"
ATTACHMENT_SAVE_PATH = os.getenv("ATTACHMENT_SAVE_PATH", r"C:\BBKM_InvoiceSorter\Invoices")

# Categorisation
KEYWORDS = [
    "Service Agreement", "Reminder", "quote", "Remittance",
    "Statement", "Caution Email", "Credit Adj"
]
CATEGORY_NAME_INFO = "Info"
CATEGORY_NAME_ATTACHMENT = "Attachment Extracted"
CATEGORY_NAME_FLAGGED = "Flagged"

# Attachment filtering
IGNORE_EXTENSIONS = {"png", "jpg", "jpeg", "gif", "bmp", "tiff", "ico"}
MIN_FILE_SIZE = 100_000  # 100KB

# Basic guard
def _require_env(var: str):
    val = os.getenv(var)
    if not val:
        raise RuntimeError(
            f"Missing required environment variable: {var}. "
            f"Set it via system env or a local .env file."
        )
    return val

# Ensure required env vars exist
_require_env("AZURE_TENANT_ID")
_require_env("AZURE_CLIENT_ID")
_require_env("AZURE_CLIENT_SECRET")


# =========================
# HTTP session with retries
# =========================
def make_session() -> requests.Session:
    s = requests.Session()
    retries = Retry(
        total=5,
        connect=5,
        read=5,
        backoff_factor=0.5,
        status_forcelist=(429, 500, 502, 503, 504),
        allowed_methods=frozenset(["GET", "POST", "PATCH"])
    )
    adapter = HTTPAdapter(max_retries=retries, pool_connections=10, pool_maxsize=10)
    s.mount("https://", adapter)
    s.mount("http://", adapter)
    return s


# =========================
# Auth
# =========================
def get_access_token() -> str:
    authority = f"https://login.microsoftonline.com/{TENANT_ID}"
    app = ConfidentialClientApplication(
        CLIENT_ID,
        authority=authority,
        client_credential=CLIENT_SECRET
    )
    token = app.acquire_token_for_client(GRAPH_SCOPE)
    if "access_token" not in token:
        raise RuntimeError(f"Failed to get token: {token}")
    return token["access_token"]


# =========================
# Graph helpers
# =========================
def paged_get(session: requests.Session, url: str, headers: Dict[str, str]) -> List[Dict]:
    """Follow @odata.nextLink and collect results."""
    items: List[Dict] = []
    while url:
        resp = session.get(url, headers=headers, timeout=60)
        if resp.status_code != 200:
            raise RuntimeError(f"GET {url} -> {resp.status_code}: {resp.text[:500]}")
        data = resp.json()
        values = data.get("value", [])
        items.extend(values)
        url = data.get("@odata.nextLink")
    return items


def get_inbox_messages(session: requests.Session, access_token: str, top: int = 50) -> List[Dict]:
    url = f"{GRAPH_BASE}/users/{USER_EMAIL}/mailFolders/inbox/messages?$top={top}&$select=id,subject,categories,hasAttachments"
    headers = {"Authorization": f"Bearer {access_token}"}
    return paged_get(session, url, headers)


def get_message(session: requests.Session, access_token: str, message_id: str) -> Dict:
    url = f"{GRAPH_BASE}/users/{USER_EMAIL}/messages/{message_id}"
    headers = {"Authorization": f"Bearer {access_token}"}
    resp = session.get(url, headers=headers, timeout=60)
    if resp.status_code != 200:
        raise RuntimeError(f"GET message {message_id} -> {resp.status_code}: {resp.text[:500]}")
    return resp.json()


def get_attachments(session: requests.Session, access_token: str, message_id: str) -> List[Dict]:
    url = f"{GRAPH_BASE}/users/{USER_EMAIL}/messages/{message_id}/attachments?$top=50"
    headers = {"Authorization": f"Bearer {access_token}"}
    return paged_get(session, url, headers)


def update_categories(session: requests.Session, access_token: str, message_id: str, add_label: str) -> None:
    # Merge new category with existing
    msg = get_message(session, access_token, message_id)
    current = set((msg.get("categories") or []))
    if add_label not in current:
        current.add(add_label)
        url = f"{GRAPH_BASE}/users/{USER_EMAIL}/messages/{message_id}"
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        }
        body = json.dumps({"categories": sorted(current)})
        r = session.patch(url, headers=headers, data=body, timeout=60)
        if r.status_code not in (200, 202):
            raise RuntimeError(f"PATCH categories -> {r.status_code}: {r.text[:500]}")


# =========================
# Attachment save
# =========================
def ensure_dir(path: str) -> None:
    if not os.path.exists(path):
        os.makedirs(path, exist_ok=True)


def is_inline_or_tiny(att: Dict) -> bool:
    # Inline images (signatures) often have isInline true or small size with image extensions.
    ext = (att.get("name") or "").split(".")[-1].lower()
    size = int(att.get("size") or 0)
    is_inline = bool(att.get("isInline"))
    if is_inline:
        return True
    if ext in IGNORE_EXTENSIONS and size < MIN_FILE_SIZE:
        return True
    return False


def save_file_attachment(att: Dict, save_dir: str) -> Optional[str]:
    """
    Save file attachment. Supports:
    - fileAttachment returned with 'contentBytes'
    - fallback to download stream if needed (rare)
    """
    ensure_dir(save_dir)

    name = att.get("name") or "attachment"
    if is_inline_or_tiny(att):
        print(f"🛑 Skipping inline/tiny attachment: {name} ({att.get('size','?')} bytes)")
        return None

    # Only handle file attachments
    odata_type = att.get("@odata.type", "").lower()
    if "fileattachment" not in odata_type:
        # Skip item attachments (embedded emails/events) for this workflow
        print(f"ℹ️ Skipping non-file attachment: {name} ({odata_type})")
        return None

    content_bytes = att.get("contentBytes")
    file_path = os.path.join(save_dir, name)

    if content_bytes:
        with open(file_path, "wb") as f:
            f.write(base64.b64decode(content_bytes))
        print(f"✔ Saved attachment: {file_path}")
        return file_path

    # Very rare: fallback stream download (if no contentBytes present)
    content_id = att.get("id")
    if not content_id:
        print(f"⚠️ No content for attachment: {name}")
        return None

    # NOTE: Graph supports /attachments/{id}/$value for raw. Kept here for completeness.
    # This requires an additional GET; some tenants may restrict it.
    print(f"⬇️ Attempting raw download for: {name}")
    return None  # Omit raw fetch to keep scope simple/safe for now.


# =========================
# Core processing
# =========================
def subject_has_keyword(subject: str) -> bool:
    s = (subject or "").lower()
    return any(k.lower() in s for k in KEYWORDS)


def process_emails() -> None:
    access_token = get_access_token()
    session = make_session()

    messages = get_inbox_messages(session, access_token, top=50)
    if not messages:
        print("No emails retrieved.")
        return

    ensure_dir(ATTACHMENT_SAVE_PATH)

    print(f"Processing {len(messages)} emails...")
    for msg in messages:
        msg_id = msg["id"]
        subject = (msg.get("subject") or "")
        clean_subject = (subject[:50] + "...") if len(subject) > 50 else subject
        categories = msg.get("categories") or []

        print(f"— Checking: {clean_subject}")

        # Skip if already categorized by our labels
        if set(categories) & {CATEGORY_NAME_INFO, CATEGORY_NAME_ATTACHMENT, CATEGORY_NAME_FLAGGED}:
            print(f"  ↪ Skipping already-categorized email")
            continue

        # Categorise by keyword OR extract attachments
        if subject_has_keyword(subject):
            print(f"  ✔ Keyword match → add category '{CATEGORY_NAME_INFO}'")
            update_categories(session, access_token, msg_id, CATEGORY_NAME_INFO)
            continue

        # Extract attachments
        attachments = get_attachments(session, access_token, msg_id)
        if not attachments:
            print(f"  (No attachments) → add category '{CATEGORY_NAME_FLAGGED}'")
            update_categories(session, access_token, msg_id, CATEGORY_NAME_FLAGGED)
            continue

        any_saved = False
        for att in attachments:
            saved = save_file_attachment(att, ATTACHMENT_SAVE_PATH)
            if saved:
                any_saved = True

        if any_saved:
            print(f"  ✔ Saved one or more attachments → add category '{CATEGORY_NAME_ATTACHMENT}'")
            update_categories(session, access_token, msg_id, CATEGORY_NAME_ATTACHMENT)
        else:
            print(f"  🚩 Nothing saved → add category '{CATEGORY_NAME_FLAGGED}'")
            update_categories(session, access_token, msg_id, CATEGORY_NAME_FLAGGED)


if __name__ == "__main__":
    try:
        process_emails()
    except Exception as e:
        print(f"ERROR: {e}")
        # Gentle backoff if you wire this into a loop elsewhere
        time.sleep(2)
