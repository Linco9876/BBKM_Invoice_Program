import os
import re
import pandas as pd
import shutil
import tempfile
import time
import sys
import spacy
from PyPDF2 import PdfReader
from pdf2image import convert_from_path
from pytesseract import image_to_string, pytesseract
from PIL import Image, ImageEnhance
from typing import Dict, Optional, Tuple

from save_attachments_from_outlook_folder import (
    GraphEmailProxy,
    _find_mail_folder_id,
    _get_access_token as _graph_get_access_token,
    _make_session as _graph_make_session,
    _move_message,
)

pytesseract.tesseract_cmd = r"C:\BBKM_InvoiceSorter\Library\Tesseract-OCR\tesseract.exe"

# Enable verbose logging
VERBOSE_LOGGING = True

def verbose_log(message):
    if VERBOSE_LOGGING:
        print(message)

def clean_name(name):
    cleaned = re.sub(r'[^\w\s]', '', name).strip()
    return re.sub(r'[-\']', ' ', cleaned)

nlp = spacy.load("en_core_web_sm")

def find_name_match(name, text, proximity=5):
    # Clean and prepare the name and text
    name = re.sub(r'[^\w\s]', '', name).strip().replace(':', '').replace('NDIS', '')
    text = re.sub(r'[^\w\s]', '', text).strip().replace(':', '').replace('NDIS', '')

    # Handle the case where the first and last name are connected by a comma with no space
    text = text.replace(',', '')

    # Split the name into parts, ignoring titles
    name_parts = [part.lower() for part in name.split() if part.lower() not in ['mr', 'mrs', 'miss', 'master', 'NDIS']]
    
    # Generate the reversed name (last name first)
    reversed_name_parts = name_parts[::-1]

    doc = nlp(text)
    text_parts = [token.text.lower() for token in doc]

    # Check for an exact match first
    exact_match_pattern = r'\b' + ' '.join(name_parts) + r'\b'
    reversed_match_pattern = r'\b' + ' '.join(reversed_name_parts) + r'\b'

    if re.search(exact_match_pattern, text, re.IGNORECASE):
        return True

    if re.search(reversed_match_pattern, text, re.IGNORECASE):
        return True

    # Check if the name is in "Lastname, Firstname" format
    if len(reversed_name_parts) >= 2:  # Ensure there are at least two elements in reversed_name_parts
        lastname_firstname_pattern = r'\b' + reversed_name_parts[0] + r', ' + reversed_name_parts[1] + r'\b'
        if re.search(lastname_firstname_pattern, text, re.IGNORECASE):
            return True

    # Check for the "FirstnameLastnameNumbers" format
    if len(name_parts) >= 2:  # Ensure there are at least two elements in name_parts
        firstname_lastname_numbers_pattern = r'\b' + name_parts[0] + name_parts[1] + r'\d*\b'
        if re.search(firstname_lastname_numbers_pattern, text, re.IGNORECASE):
            return True

    # Check if the name parts are close together (both normal and reversed)
    for i in range(len(text_parts)):
        if len(name_parts) > 0 and name_parts[0] == text_parts[i]:  # Ensure name_parts is not empty
            for j in range(1, len(name_parts)):
                if i + j < len(text_parts) and name_parts[j] == text_parts[i + j]:
                    if j <= proximity:
                        return True
                    else:
                        break

        if len(reversed_name_parts) > 0 and reversed_name_parts[0] == text_parts[i]:  # Ensure reversed_name_parts is not empty
            for j in range(1, len(reversed_name_parts)):
                if i + j < len(text_parts) and reversed_name_parts[j] == text_parts[i + j]:
                    if j <= proximity:
                        return True
                    else:
                        break

    return False

_GRAPH_FOLDER_CACHE: Dict[str, str] = {}


def _ensure_graph_handles(
    email: Optional[GraphEmailProxy],
) -> Tuple[object, str]:
    session = getattr(email, "_session", None) if email else None
    token = getattr(email, "_access_token", None) if email else None

    if session and token:
        return session, token

    session = _graph_make_session()
    token = _graph_get_access_token()

    if email:
        setattr(email, "_session", session)
        setattr(email, "_access_token", token)

    return session, token


def move_email(email, subfolder_name, filename):
    try:
        if email and not email.IsConflict:
            session, token = _ensure_graph_handles(email)
            cache_key = subfolder_name.strip().casefold()
            folder_id = _GRAPH_FOLDER_CACHE.get(cache_key)
            if not folder_id:
                folder_id = _find_mail_folder_id(
                    session, token, display_name=subfolder_name
                )
                _GRAPH_FOLDER_CACHE[cache_key] = folder_id

            _move_message(session, token, email.id, folder_id)
        else:
            print(f"Email.IsConflict is True for file: {filename}")
    except Exception as e:
        print(f"Error moving email: {e}")

def get_subfolder(code, subfolder_paths):
    return subfolder_paths

def find_name_code_match(text, excel_data):
    for i in range(len(excel_data)):
        name = str(excel_data.iloc[i, 0])
        code = excel_data.iloc[i, 1]

        if find_name_match(name, text):
            return True, code

    return False, None


def find_existing_code_prefix(filename, excel_data):
    base_name = os.path.splitext(filename)[0].lower()

    for i in range(len(excel_data)):
        raw_code = excel_data.iloc[i, 1]
        if pd.isna(raw_code):
            continue

        code = str(raw_code).strip()
        if not code:
            continue

        code_lower = code.lower()
        normalized_code = re.sub(r"[^a-z0-9]", "", code_lower)

        if base_name.startswith(code_lower):
            return code

        for separator in ("_", "-", " "):
            if base_name.startswith(f"{code_lower}{separator}"):
                return code

        normalized_base = re.sub(r"[^a-z0-9]", "", base_name)
        if normalized_code and normalized_base.startswith(normalized_code):
            return code

    return None

def handle_successful_match(filename, file_path, code, renamed_invoices_path, failed_path, email_file_map, method):
    new_filename = f"{code}_{filename}"
    target_folder = get_subfolder(code, renamed_invoices_path)
    new_file_path = os.path.join(target_folder, new_filename)

    if os.path.exists(new_file_path):
        handle_doubled_up(filename, file_path, failed_path, email_file_map)
    else:
        move_file_and_update_email(filename, file_path, new_file_path, "Complete invoices", email_file_map)

    if method == 'PyPDF2':
        print(f"PDF match found")
    elif method == 'pytesseract':
        print(f"OCR match found")


def handle_existing_code_match(filename, file_path, code, renamed_invoices_path, failed_path, email_file_map):
    target_folder = get_subfolder(code, renamed_invoices_path)
    destination_path = os.path.join(target_folder, filename)

    if os.path.exists(destination_path):
        handle_doubled_up(filename, file_path, failed_path, email_file_map)
        return

    move_file_and_update_email(
        filename,
        file_path,
        destination_path,
        "Complete invoices",
        email_file_map,
    )
    print(f"Existing code detected - moved {filename}")


def filename_has_code_prefix(filename: str, code: str) -> bool:
    """Return True if the filename already starts with the given code."""

    if not code:
        return False

    base_name = os.path.splitext(filename)[0].lower()
    code_lower = str(code).strip().lower()

    # Direct prefix checks (e.g. "abc123" or "abc123_")
    if base_name.startswith(code_lower):
        return True

    for separator in ("_", "-", " "):
        if base_name.startswith(f"{code_lower}{separator}"):
            return True

    # Normalized check to allow codes with special characters
    normalized_code = re.sub(r"[^a-z0-9]", "", code_lower)
    normalized_base = re.sub(r"[^a-z0-9]", "", base_name)
    return bool(normalized_code) and normalized_base.startswith(normalized_code)

def handle_doubled_up(filename, file_path, failed_path, email_file_map):
    print(f"You've done {filename} already silly")
    email = email_file_map.get(filename)

    if email:
        email.Categories = "Doubled Up"
        email.Save()

        doubled_up_filename = f"{os.path.splitext(filename)[0]}_Doubled_up{os.path.splitext(filename)[1]}"
        doubled_up_file_path = os.path.join(failed_path, doubled_up_filename)
        shutil.move(file_path, doubled_up_file_path)

        move_email(email, "Complete invoices", filename)

def move_file_and_update_email(filename, file_path, new_file_path, target_folder_name, email_file_map):
    shutil.move(file_path, new_file_path)
    print(f"Success {os.path.basename(new_file_path)}")

    email = email_file_map.get(filename)
    if email:
        move_email(email, target_folder_name, filename)

def handle_failed_file(filename, file_path, failed_path, email_file_map, text):
    failed_counter = 1
    failed_file_name = filename
    failed_file_path = os.path.join(failed_path, failed_file_name)

    while os.path.exists(failed_file_path):
        file_name, file_extension = os.path.splitext(filename)
        failed_file_name = f"{file_name}_failed{failed_counter}{file_extension}"
        failed_file_path = os.path.join(failed_path, failed_file_name)
        failed_counter += 1

    if os.path.exists(file_path):
        shutil.move(file_path, failed_file_path)
        print(f"Failed Moving {failed_file_name}")
        # print(f"Extracted text:\n{text}")

    email = email_file_map.get(filename)
    if email:
        try:
            email.Categories = "Failed Rename"
            email.Save()
            move_email(email, "Complete invoices", filename)  # Move the email to the "Complete invoices" folder
        except Exception as e:
            print(f"Error handling failed file: {e}")

def extract_text_ocr(file_path):
    try:
        Image.MAX_IMAGE_PIXELS = None 
        images = convert_from_path(file_path, dpi=300) 
        text = ""
        for img in images:
            img = ImageEnhance.Contrast(img).enhance(1.5)
            extracted_text = image_to_string(img)
            text += extracted_text
        return text
    except Exception as e:
        print(f"Error extracting text with OCR: {e}")
        return ""

def extract_text_pypdf2(file_path):
    with open(file_path, 'rb') as f:
        reader = PdfReader(f)
        text = ''
        for page in reader.pages:
            text += page.extract_text()
    return text

def process_pdf(filename, file_path, text, excel_data, renamed_invoices_path, failed_path, email_file_map, method):
    found_match, code = find_name_code_match(text, excel_data)
    if found_match:
        handle_successful_match(filename, file_path, code, renamed_invoices_path, failed_path, email_file_map, method)
        return True
    else:
        return False

def process_pdfs(pdf_files, invoices_path, excel_data, renamed_invoices_path, failed_path, email_file_map):
    for filename in pdf_files:
        file_path = os.path.join(invoices_path, filename)
        
        # First, check if the file name already contains a known client code
        code_prefix = find_existing_code_prefix(filename, excel_data)
        if code_prefix:
            handle_existing_code_match(
                filename,
                file_path,
                code_prefix,
                renamed_invoices_path,
                failed_path,
                email_file_map,
            )
            continue

        # Next, try to find a client name match directly in the file name
        found_match, code = find_name_code_match(filename, excel_data)

        if found_match:
            if filename_has_code_prefix(filename, code):
                handle_existing_code_match(
                    filename,
                    file_path,
                    code,
                    renamed_invoices_path,
                    failed_path,
                    email_file_map,
                )
            else:
                handle_successful_match(filename, file_path, code, renamed_invoices_path, failed_path, email_file_map, 'Filename')
            continue

        # If no match found in the file name, proceed with PyPDF2 extraction
        try:
            text = extract_text_pypdf2(file_path)
            found_match = process_pdf(filename, file_path, text, excel_data, renamed_invoices_path, failed_path, email_file_map, 'PyPDF2')
        except Exception as e:
            found_match = False

        # If no match found using PyPDF2, proceed with OCR extraction
        if not found_match:
            text = extract_text_ocr(file_path)
            found_match = process_pdf(filename, file_path, text, excel_data, renamed_invoices_path, failed_path, email_file_map, 'pytesseract')

        if not found_match:
            handle_failed_file(filename, file_path, failed_path, email_file_map, text)

def read_csv_data(csv_file):
    with tempfile.NamedTemporaryFile(delete=False) as temp_file:
        shutil.copy2(csv_file, temp_file.name)
        try:
            csv_data = pd.read_csv(temp_file.name, encoding='utf-8', on_bad_lines='skip')
        except UnicodeDecodeError:
            csv_data = pd.read_csv(temp_file.name, encoding='ISO-8859-1', on_bad_lines='skip')
    os.unlink(temp_file.name)
    return csv_data

def pytesseract_main(updated_saved_attachments, email_file_map):
    invoice_path = r"C:\BBKM_InvoiceSorter\Invoices"
    csv_file = r"C:\Users\Administrator\Better Bookkeeping Management\BBKM - Documents\BBKM Plan Management\Client Names.csv"
    renamed_invoices_path = os.path.join(invoice_path, "Renamed Invoices")
    failed_path = os.path.join(invoice_path, "Failed")

    os.makedirs(renamed_invoices_path, exist_ok=True)
    os.makedirs(failed_path, exist_ok=True)

    try:
        csv_data = read_csv_data(csv_file)
    except PermissionError:
        print("Someone opened the CSV file. Waiting for 3 minutes before retrying...")
        time.sleep(180)
        print("Restarting the script...")
        script_path = f'"{sys.argv[0]}"'
        os.execl(sys.executable, sys.executable, script_path, *sys.argv[1:])
        return

    email_file_map_copy = email_file_map.copy()
    for file_path, email in zip(updated_saved_attachments, email_file_map_copy.values()):
        email_file_map[os.path.basename(file_path)] = email

    pdf_files = [f for f in os.listdir(invoice_path) if f.lower().endswith('.pdf')]
    process_pdfs(pdf_files, invoice_path, csv_data, renamed_invoices_path, failed_path, email_file_map)


def main():
    # Load the excel data
    excel_data = pd.read_csv(r'C:\Users\Administrator\Better Bookkeeping Management\BBKM - Documents\BBKM Plan Management\Client Names.CSV', header=None)

    # Define folder paths
    invoices_path = 'C:/BBKM_InvoiceSorter/Invoices'
    renamed_invoices_path = 'C:/BBKM_InvoiceSorter/Invoices/Renamed Invoices'
    failed_path = 'C:/BBKM_InvoiceSorter/Invoices/Failed'

    # Get the email_file_map
    email_file_map = {}  # You will need to implement a function to create this mapping based on your needs.

    while True:
        # Get the PDF files from the invoices folder
        pdf_files = [f for f in os.listdir(invoices_path) if f.lower().endswith('.pdf')]

        # Process the PDF files
        process_pdfs(pdf_files, invoices_path, excel_data, renamed_invoices_path, failed_path, email_file_map)

        pytesseract_main(pdf_files, email_file_map)

if __name__ == "__main__":
    main()
