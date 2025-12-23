import os
import shutil
import re
import time
import hashlib
import tempfile
import pandas as pd
from pytesseract import pytesseract
from pdf2image import convert_from_path
from PIL import Image, ImageFilter, ImageOps
import sqlite3
from datetime import datetime, timedelta

# -------------------- Config & Paths --------------------
pytesseract.tesseract_cmd = r"C:\BBKM_InvoiceSorter\Library\Tesseract-OCR\tesseract.exe"

NDIS_STATEMENT_PATH = r"C:\Users\Administrator\Better Bookkeeping Management\BBKM - Documents\BBKM Plan Management\NDIS\ZInvoices for lodgement\Invoice Program\NDIS Activity Statement"
SRC_FOLDER = r"C:\BBKM_InvoiceSorter\Invoices\Renamed Invoices"
DEST_FOLDER = r"C:\Users\Administrator\Better Bookkeeping Management\BBKM - Documents\BBKM Plan Management\NDIS\ZInvoices for lodgement\Invoice Program"
RECEIPTS_FOLDER = os.path.join(DEST_FOLDER, "Renamed Receipts")
MANUAL_LODGEMENT_FOLDER = os.path.join(DEST_FOLDER, "Manual Lodgement")
NEW_PROVIDER_FOLDER = os.path.join(DEST_FOLDER, "New Provider")
UNASSIGNED_PLAN_MANAGER_FOLDER = os.path.join(DEST_FOLDER, "Unassigned Plan Manager")
SRC_FOLDER_ATTEMPT = os.path.join(DEST_FOLDER, "Attempt Code")
DEST_FOLDER_ATTEMPT = r"C:\BBKM_InvoiceSorter\Invoices"
SRC_FOLDER_FAILED = r"C:\BBKM_InvoiceSorter\Invoices\Failed"
DEST_FOLDER_FAILED = os.path.join(DEST_FOLDER, "Failed to Code")
STA_INVOICES_FOLDER = os.path.join(DEST_FOLDER, "STA and Assistance")
STREAMLINE_FOLDER = os.path.join(DEST_FOLDER, "Streamline Invoices")
AT_CONSUMABLES_FOLDER = os.path.join(DEST_FOLDER, "AT&Consumables")

# Failed subfolders
STREAMLINE_FAILED_FOLDER = os.path.join(DEST_FOLDER_FAILED, "Streamline failed to code")
FAILED_MANUAL_FOLDER = os.path.join(DEST_FOLDER_FAILED, "Failed Manual Lodgment")
FAILED_AT_FOLDER = os.path.join(DEST_FOLDER_FAILED, "Failed AT&Consumables")

# Vendor map & logs
VENDOR_CSV_PATH = r"C:\BBKM_InvoiceSorter\Scripts\Vendors.csv"
MISSING_FILES_LOG = r"C:\BBKM_InvoiceSorter\missing_files.log"
CLIENT_PROFILES_PATH = r"C:\Users\Administrator\Better Bookkeeping Management\BBKM - Documents\BBKM Plan Management\Client_Profiles.csv"

# Quarantine for final fallback when moves keep failing
COULD_NOT_MOVE_FOLDER = os.path.join(DEST_FOLDER_FAILED, "Could not move")
os.makedirs(UNASSIGNED_PLAN_MANAGER_FOLDER, exist_ok=True)
os.makedirs(COULD_NOT_MOVE_FOLDER, exist_ok=True)

# SQLite DB for 90-day duplicate detection
DB_PATH = r"C:\BBKM_InvoiceSorter\file_history.sqlite"

# -------------------- Load Vendors --------------------
def load_vendors():
    df = pd.read_csv(VENDOR_CSV_PATH)
    return df.set_index('Vendor')['FolderType'].to_dict()

VENDORS = load_vendors()


def _normalize_client_profile_columns(df: pd.DataFrame) -> pd.DataFrame:
    expected = [
        "Client Code",
        "All Known Names",
        "NDIS Number",
        "Assigned Plan Manager",
    ]

    trimmed_columns = [str(col).strip() for col in df.columns]
    df.columns = trimmed_columns

    rename_map = {}
    for idx, expected_name in enumerate(expected):
        if expected_name not in df.columns and idx < len(df.columns):
            rename_map[df.columns[idx]] = expected_name
    if rename_map:
        df = df.rename(columns=rename_map)

    for col in expected:
        if col not in df.columns:
            df[col] = None

    return df[expected]


def _load_client_profiles():
    if not os.path.exists(CLIENT_PROFILES_PATH):
        return pd.DataFrame(columns=["Client Code", "All Known Names", "NDIS Number", "Assigned Plan Manager"])

    read_kwargs = {
        "dtype": str,
        "on_bad_lines": "skip",
        "na_filter": False,
    }

    with open(CLIENT_PROFILES_PATH, "rb") as src:
        with tempfile.NamedTemporaryFile(delete=False) as temp_file:
            temp_file.write(src.read())
            temp_path = temp_file.name

    try:
        try:
            df = pd.read_csv(temp_path, encoding="utf-8", **read_kwargs)
        except UnicodeDecodeError:
            df = pd.read_csv(temp_path, encoding="ISO-8859-1", **read_kwargs)
    finally:
        os.unlink(temp_path)

    return _normalize_client_profile_columns(df)


def _build_client_records(df: pd.DataFrame):
    records = []
    for _, row in df.iterrows():
        code = str(row.get("Client Code", "") or "").strip()
        plan_manager = str(row.get("Assigned Plan Manager", "") or "").strip()
        if not code:
            continue
        records.append(
            {
                "code": code,
                "code_normalized": re.sub(r"[^a-z0-9]", "", code.lower()),
                "plan_manager": plan_manager,
            }
        )
    return records


CLIENT_PROFILES = _load_client_profiles()
CLIENT_RECORDS = _build_client_records(CLIENT_PROFILES)


def _sanitize_folder_name(name: str) -> str:
    return re.sub(r'[<>:"/\\|?*]', "_", name).strip()


def _plan_manager_root(plan_manager: str) -> str:
    label = plan_manager.strip() if plan_manager else ""
    if not label:
        return UNASSIGNED_PLAN_MANAGER_FOLDER

    safe_name = _sanitize_folder_name(label)
    return os.path.join(DEST_FOLDER, safe_name) if safe_name else UNASSIGNED_PLAN_MANAGER_FOLDER


def _lookup_plan_manager(filename: str) -> str:
    base_name = os.path.splitext(filename)[0].lower()
    normalized_base = re.sub(r"[^a-z0-9]", "", base_name)

    for record in CLIENT_RECORDS:
        code_lower = record["code"].lower()
        if base_name.startswith(code_lower):
            return record["plan_manager"]
        for separator in ("_", "-", " "):
            if base_name.startswith(f"{code_lower}{separator}"):
                return record["plan_manager"]
        if record["code_normalized"] and normalized_base.startswith(record["code_normalized"]):
            return record["plan_manager"]

    return ""

# -------------------- State Tracking --------------------
missing_files = set()
if os.path.exists(MISSING_FILES_LOG):
    with open(MISSING_FILES_LOG, 'r') as f:
        missing_files.update(line.strip() for line in f if line.strip())

# -------------------- Helpers --------------------
def preprocess_image(image):
    image = image.convert('L')
    image = ImageOps.invert(image)
    image = image.filter(ImageFilter.MedianFilter())
    image = ImageOps.autocontrast(image)
    return image

def _file_size(path: str) -> int:
    try:
        return os.path.getsize(path)
    except Exception:
        return -1

def _file_hash(path: str, block_size: int = 1024 * 1024) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for block in iter(lambda: f.read(block_size), b""):
            h.update(block)
    return h.hexdigest()

def _unique_with_counter(dest: str) -> str:
    base, ext = os.path.splitext(dest)
    i = 1
    cand = dest
    while os.path.exists(cand):
        cand = f"{base} ({i}){ext}"
        i += 1
    return cand

def _wait_for_settle(path: str, checks: int = 3, delay: float = 0.8) -> bool:
    """
    Wait until a file's size stays unchanged for `checks` intervals.
    Helps avoid processing while OneDrive is syncing or a write is in progress.
    """
    try:
        last = -1
        stable = 0
        for _ in range(checks * 3):  # give a little extra runway
            size = os.path.getsize(path)
            if size == last:
                stable += 1
                if stable >= checks:
                    return True
            else:
                stable = 0
            last = size
            time.sleep(delay)
    except Exception:
        return False
    return True

# NEW: find another file with the SAME NAME and SAME HASH in the same folder
def _find_same_name_and_hash_in_folder(folder: str, my_path: str, my_name: str, my_hash: str):
    """
    Returns the path of a sibling file that has the same name (case-insensitive)
    and identical content hash. Returns None if not found.
    """
    my_name_lower = my_name.lower()
    for name in os.listdir(folder):
        p = os.path.join(folder, name)
        if p == my_path or not os.path.isfile(p):
            continue
        if name.lower() != my_name_lower:
            continue  # must be same name
        try:
            if _file_hash(p) == my_hash:
                return p
        except Exception:
            pass
    return None

# -------------------- SQLite (SQL-only duplicate policy) --------------------
def _db_init():
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS file_history (
            sha256 TEXT PRIMARY KEY,
            size INTEGER,
            first_seen_utc INTEGER,
            last_seen_utc INTEGER,
            last_path TEXT
        )
    """)
    con.commit()
    con.close()

def _db_seen_recently(file_hash: str, days: int = 90) -> bool:
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    cutoff = int((datetime.utcnow() - timedelta(days=days)).timestamp())
    cur.execute("SELECT 1 FROM file_history WHERE sha256=? AND last_seen_utc>?", (file_hash, cutoff))
    row = cur.fetchone()
    con.close()
    return row is not None

def _db_record(path: str, file_hash: str):
    try:
        size = os.path.getsize(path)
    except Exception:
        size = None
    now = int(datetime.utcnow().timestamp())
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    cur.execute("""
        INSERT INTO file_history (sha256, size, first_seen_utc, last_seen_utc, last_path)
        VALUES (?, ?, ?, ?, ?)
        ON CONFLICT(sha256) DO UPDATE SET
            last_seen_utc=excluded.last_seen_utc,
            last_path=excluded.last_path,
            size=COALESCE(excluded.size, size)
    """, (file_hash, size, now, now, path))
    con.commit()
    con.close()

_db_init()

# -------------------- Safe Move (with quarantine fallback) --------------------
def safe_move(src, dest, log_message, retries: int = 3, wait_secs: float = 2.0):
    """
    - If dest doesn't exist → move
    - If dest exists → treat as name collision and save as '(1)', '(2)', etc.
    - If all retries fail → quarantine into 'Failed to Code/Could not move'
    """
    try:
        if not os.path.exists(src):
            # Don't spam the missing log here; upstream handles it
            print(f"File not found: {src}")
            return False

        os.makedirs(os.path.dirname(dest), exist_ok=True)

        attempt = 0
        while attempt <= retries:
            try:
                if not os.path.exists(dest):
                    shutil.move(src, dest)
                    print(log_message)
                    return True
                else:
                    # Name collision → append counter
                    dest2 = _unique_with_counter(dest)
                    shutil.move(src, dest2)
                    print(f"{log_message} (name collision → saved as '{os.path.basename(dest2)}')")
                    return True
            except Exception as e:
                attempt += 1
                print(f"Move attempt {attempt} failed for {src} → {dest}: {e}")
                if attempt > retries:
                    # Final fallback so nothing is left behind
                    base = os.path.basename(src)
                    quarantine_dest = os.path.join(COULD_NOT_MOVE_FOLDER, f"unmoved_{base}")
                    quarantine_dest = _unique_with_counter(quarantine_dest)
                    try:
                        shutil.move(src, quarantine_dest)
                        print(f"Quarantined (could not move): {src} → {quarantine_dest}")
                    except Exception as e2:
                        print(f"Failed to quarantine {src}: {e2}")
                    return False
                time.sleep(wait_secs)
    except Exception as e:
        print(f"Error in safe_move({src}): {e}")
        return False

# -------------------- Core Logic --------------------
def move_files(src_folder, dest_folder):
    try:
        files = [f for f in os.listdir(src_folder)
                 if os.path.isfile(os.path.join(src_folder, f))
                 and f.lower() != "desktop.ini"]
    except FileNotFoundError:
        print(f"Source folder not found: {src_folder}")
        return

    for filename in files:
        file_path = os.path.join(src_folder, filename)

        # Skip missing
        if not os.path.exists(file_path):
            continue

        # Wait for file to settle (avoid mid-sync/mid-write)
        if not _wait_for_settle(file_path, checks=3, delay=0.8):
            print(f"Skipping (not settled yet): {file_path}")
            continue

        # Compute content hash early for SQL duplicate policy
        try:
            file_hash = _file_hash(file_path)
        except Exception as e:
            print(f"Error hashing {file_path}: {e}")
            file_hash = None

        # STRICT double_ policy:
        # Only add 'double_' if BOTH:
        # 1) SQL says we've seen this hash within 90 days
        # 2) Another file in this source folder has the SAME NAME and SAME HASH
        if file_hash and _db_seen_recently(file_hash, 90):
            match_path = _find_same_name_and_hash_in_folder(src_folder, file_path, filename, file_hash)
            if match_path:
                new_filename = f"double_{filename}"
                new_path = os.path.join(src_folder, new_filename)
                if os.path.exists(new_path):
                    new_path = _unique_with_counter(new_path)
                try:
                    os.rename(file_path, new_path)
                    print(
                        f"Duplicate present (same name & same content in source): "
                        f"matched='{match_path}' → renamed '{filename}' → '{os.path.basename(new_path)}'"
                    )
                    filename = os.path.basename(new_path)
                    file_path = new_path
                except Exception as e:
                    print(f"Failed to prefix 'double_' for {file_path}: {e}")

        plan_manager = _lookup_plan_manager(filename)
        plan_manager_base = _plan_manager_root(plan_manager)

        # Attempt Code short-circuit
        if src_folder == SRC_FOLDER_ATTEMPT:
            dest_path = os.path.join(dest_folder, filename)
            if safe_move(file_path, dest_path, f"Moved from Attempt Code -> Invoices: {file_path}"):
                if file_hash:
                    _db_record(dest_path, file_hash)
            continue

        try:
            # Receipts: keep original condition tied to dest folder label
            if "receipt" in filename.lower() and os.path.split(dest_folder)[1] == "Invoice Program":
                destination = os.path.join(RECEIPTS_FOLDER, filename)
                if safe_move(file_path, destination, f"Moved receipt: {file_path}"):
                    if file_hash:
                        _db_record(destination, file_hash)
                continue

            # Process PDFs from SRC/FAILED
            if filename.lower().endswith(".pdf") and src_folder in [SRC_FOLDER, SRC_FOLDER_FAILED]:
                # Verify readable PDF (render) else mark corrupt
                try:
                    images = convert_from_path(file_path)
                except Exception:
                    corrupt_filename = f"corrupt_{filename}"
                    corrupt_dest = os.path.join(DEST_FOLDER_FAILED, corrupt_filename)
                    if safe_move(file_path, corrupt_dest, f"Corrupt file moved: {file_path}"):
                        if file_hash:
                            _db_record(corrupt_dest, file_hash)
                    continue

                found_sta = found_respite = found_ndis_statement = False
                found_vendor = None

                # OCR to detect vendor and categories
                for image in images:
                    content = pytesseract.image_to_string(preprocess_image(image)).lower().replace(" ", "")
                    if "ndisactivitystatement" in content:
                        found_ndis_statement = True
                        break
                    if re.search(r'\bsta\b|\b\dsta\b', content):
                        found_sta = True
                    if re.search(r'inc\.\srespite', content):
                        found_respite = True
                    for vendor_raw, folder_type in VENDORS.items():
                        cleaned_vendor = vendor_raw.replace(" ", "").lower()
                        if cleaned_vendor in content:
                            found_vendor = vendor_raw
                            break

                # NDIS activity statements
                if found_ndis_statement:
                    dest_path = os.path.join(NDIS_STATEMENT_PATH, filename)
                    if safe_move(file_path, dest_path, f"NDIS statement moved: {file_path}"):
                        if file_hash:
                            _db_record(dest_path, file_hash)
                    continue

                # Files already in Failed
                if src_folder == SRC_FOLDER_FAILED:
                    if found_vendor:
                        folder_type = VENDORS.get(found_vendor)
                        if folder_type == 1:
                            target = STREAMLINE_FAILED_FOLDER
                        elif folder_type == 2:
                            target = FAILED_MANUAL_FOLDER
                        elif folder_type == 3:
                            target = FAILED_AT_FOLDER
                        else:
                            target = DEST_FOLDER_FAILED
                        log = f"Moved failed invoice to {os.path.basename(target)}: {file_path}"
                    else:
                        target = DEST_FOLDER_FAILED
                        log = f"Moved unknown failed invoice: {file_path}"
                    dest_path = os.path.join(target, filename)
                    if safe_move(file_path, dest_path, log):
                        if file_hash:
                            _db_record(dest_path, file_hash)
                    continue

                # Fallback: infer vendor from filename if OCR missed it
                if not found_vendor:
                    stem, _ext = os.path.splitext(filename)
                    norm_name = re.sub(r'[\s\u00A0._-]+', ' ', stem).strip().lower()
                    squashed = norm_name.replace(" ", "")
                    for vendor in VENDORS:
                        if vendor.replace(" ", "").lower() in squashed:
                            found_vendor = vendor
                            break

                # Choose destination
                if found_sta or found_respite:
                    target = os.path.join(plan_manager_base, os.path.basename(STA_INVOICES_FOLDER))
                    log = f"Moved to STA and Assistance: {file_path}"
                elif found_vendor:
                    folder_type = VENDORS[found_vendor]
                    if folder_type == 1:
                        target = os.path.join(plan_manager_base, os.path.basename(STREAMLINE_FOLDER))
                        log = f"Moved to Streamline: {file_path}"
                    elif folder_type == 2:
                        target = os.path.join(plan_manager_base, os.path.basename(MANUAL_LODGEMENT_FOLDER))
                        log = f"Moved to Manual Lodgement: {file_path}"
                    elif folder_type == 3:
                        target = os.path.join(plan_manager_base, os.path.basename(AT_CONSUMABLES_FOLDER))
                        log = f"Moved to AT&Consumables: {file_path}"
                    else:
                        target = os.path.join(plan_manager_base, found_vendor)
                        log = f"Moved to custom vendor ({found_vendor}) under {plan_manager or 'Unassigned'}: {file_path}"
                else:
                    target = os.path.join(plan_manager_base, os.path.basename(NEW_PROVIDER_FOLDER))
                    log = f"Moved to New Provider (no match) under {plan_manager or 'Unassigned'}: {file_path}"

                dest_path = os.path.join(target, filename)
                if safe_move(file_path, dest_path, log):
                    if file_hash:
                        _db_record(dest_path, file_hash)

        except Exception as e:
            print(f"Error processing {filename}: {e}")

# -------------------- Entrypoint --------------------
if __name__ == "__main__":
    move_files(SRC_FOLDER_ATTEMPT, DEST_FOLDER_ATTEMPT)
    move_files(SRC_FOLDER_FAILED, DEST_FOLDER_FAILED)
    move_files(SRC_FOLDER, DEST_FOLDER)
