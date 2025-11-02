import os
import re
import sys
import time
import shutil
import tempfile
from typing import Dict, List, Optional

import pandas as pd
import requests
import spacy
from PyPDF2 import PdfReader
from pdf2image import convert_from_path
from PIL import Image, ImageEnhance
from pytesseract import image_to_string, pytesseract
import msal

pytesseract.tesseract_cmd = r"C:\BBKM_InvoiceSorter\Library\Tesseract-OCR\tesseract.exe"

GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"


class GraphEmailClient:
    """Helper for interacting with Microsoft Graph for mailbox operations."""

    def __init__(
        self,
        tenant_id: str,
        client_id: str,
        client_secret: str,
        mailbox: str,
        scope: Optional[List[str]] = None,
    ) -> None:
        authority = f"https://login.microsoftonline.com/{tenant_id}"
        self._scope = scope or ["https://graph.microsoft.com/.default"]
        self._mailbox = mailbox
        self._folder_cache: Dict[str, str] = {}
        self._session = requests.Session()
        self._app = msal.ConfidentialClientApplication(
            client_id=client_id,
            authority=authority,
            client_credential=client_secret,
        )

    def _acquire_token(self) -> str:
        result = self._app.acquire_token_silent(self._scope, account=None)
        if not result:
            result = self._app.acquire_token_for_client(scopes=self._scope)

        if "access_token" not in result:
            raise RuntimeError(f"Failed to acquire Graph token: {result.get('error_description', 'Unknown error')}")

        return result["access_token"]

    def _headers(self) -> Dict[str, str]:
        return {
            "Authorization": f"Bearer {self._acquire_token()}",
            "Content-Type": "application/json",
        }

    def _normalize_key(self, value: str) -> str:
        return value.strip().lower()

    def get_folder_id(self, display_name: str) -> str:
        cache_key = self._normalize_key(display_name)
        if cache_key in self._folder_cache:
            return self._folder_cache[cache_key]

        filter_value = display_name.replace("'", "''")
        params = {"$filter": f"displayName eq '{filter_value}'", "$top": "1"}
        url = f"{GRAPH_BASE_URL}/users/{self._mailbox}/mailFolders"
        response = self._session.get(url, headers=self._headers(), params=params)
        if response.status_code != 200:
            raise RuntimeError(
                f"Unable to resolve folder '{display_name}': {response.status_code} {response.text}"
            )

        data = response.json()
        folders = data.get("value", [])
        if not folders:
            raise ValueError(f"Folder '{display_name}' not found in mailbox {self._mailbox}.")

        folder_id = folders[0]["id"]
        self._folder_cache[cache_key] = folder_id
        return folder_id

    def move_message(self, message_id: str, destination_folder: str) -> None:
        folder_id = self.get_folder_id(destination_folder)
        url = f"{GRAPH_BASE_URL}/users/{self._mailbox}/messages/{message_id}/move"
        payload = {"destinationId": folder_id}
        response = self._session.post(url, headers=self._headers(), json=payload)
        if response.status_code not in (200, 201):
            raise RuntimeError(
                f"Failed to move message {message_id} to '{destination_folder}': {response.status_code} {response.text}"
            )

    def update_categories(self, message_id: str, categories: List[str]) -> None:
        url = f"{GRAPH_BASE_URL}/users/{self._mailbox}/messages/{message_id}"
        payload = {"categories": categories}
        response = self._session.patch(url, headers=self._headers(), json=payload)
        if response.status_code not in (200, 202):
            raise RuntimeError(
                f"Failed to update categories for message {message_id}: {response.status_code} {response.text}"
            )


def create_graph_email_client() -> GraphEmailClient:
    tenant_id = os.environ.get("GRAPH_TENANT_ID")
    client_id = os.environ.get("GRAPH_CLIENT_ID")
    client_secret = os.environ.get("GRAPH_CLIENT_SECRET")
    mailbox = os.environ.get("GRAPH_MAILBOX")

    missing = [
        name
        for name, value in [
            ("GRAPH_TENANT_ID", tenant_id),
            ("GRAPH_CLIENT_ID", client_id),
            ("GRAPH_CLIENT_SECRET", client_secret),
            ("GRAPH_MAILBOX", mailbox),
        ]
        if not value
    ]

    if missing:
        raise EnvironmentError(
            "Missing required Graph configuration environment variables: " + ", ".join(missing)
        )

    return GraphEmailClient(
        tenant_id=tenant_id,
        client_id=client_id,
        client_secret=client_secret,
        mailbox=mailbox,
    )

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

def move_email(message_id: Optional[str], subfolder_name: str, filename: str, graph_client: GraphEmailClient) -> None:
    if not message_id:
        print(f"No message ID available to move email for file: {filename}")
        return

    try:
        graph_client.move_message(message_id, subfolder_name)
    except Exception as exc:
        print(f"Error moving email for {filename}: {exc}")

def get_subfolder(code, subfolder_paths):
    return subfolder_paths

def find_name_code_match(text, excel_data):
    for i in range(len(excel_data)):
        name = str(excel_data.iloc[i, 0])
        code = excel_data.iloc[i, 1]

        if find_name_match(name, text):
            return True, code

    return False, None

def handle_successful_match(
    filename,
    file_path,
    code,
    renamed_invoices_path,
    failed_path,
    email_file_map,
    method,
    graph_client: GraphEmailClient,
):
    new_filename = f"{code}_{filename}"
    target_folder = get_subfolder(code, renamed_invoices_path)
    new_file_path = os.path.join(target_folder, new_filename)

    if os.path.exists(new_file_path):
        handle_doubled_up(filename, file_path, failed_path, email_file_map, graph_client)
    else:
        move_file_and_update_email(
            filename,
            file_path,
            new_file_path,
            "Complete invoices",
            email_file_map,
            graph_client,
        )

    if method == 'PyPDF2':
        print(f"PDF match found")
    elif method == 'pytesseract':
        print(f"OCR match found")

def handle_doubled_up(filename, file_path, failed_path, email_file_map, graph_client: GraphEmailClient):
    print(f"You've done {filename} already silly")
    message_id = email_file_map.get(filename)

    if message_id:
        try:
            graph_client.update_categories(message_id, ["Doubled Up"])
        except Exception as exc:
            print(f"Error updating categories for doubled up file {filename}: {exc}")

        doubled_up_filename = f"{os.path.splitext(filename)[0]}_Doubled_up{os.path.splitext(filename)[1]}"
        doubled_up_file_path = os.path.join(failed_path, doubled_up_filename)
        shutil.move(file_path, doubled_up_file_path)

        move_email(message_id, "Complete invoices", filename, graph_client)

def move_file_and_update_email(
    filename,
    file_path,
    new_file_path,
    target_folder_name,
    email_file_map,
    graph_client: GraphEmailClient,
):
    shutil.move(file_path, new_file_path)
    print(f"Success {os.path.basename(new_file_path)}")

    message_id = email_file_map.get(filename)
    if message_id:
        move_email(message_id, target_folder_name, filename, graph_client)

def handle_failed_file(filename, file_path, failed_path, email_file_map, text, graph_client: GraphEmailClient):
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

    message_id = email_file_map.get(filename)
    if message_id:
        try:
            graph_client.update_categories(message_id, ["Failed Rename"])
        except Exception as exc:
            print(f"Error updating categories for failed file {filename}: {exc}")
        move_email(message_id, "Complete invoices", filename, graph_client)

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

def process_pdf(
    filename,
    file_path,
    text,
    excel_data,
    renamed_invoices_path,
    failed_path,
    email_file_map,
    method,
    graph_client: GraphEmailClient,
):
    found_match, code = find_name_code_match(text, excel_data)
    if found_match:
        handle_successful_match(
            filename,
            file_path,
            code,
            renamed_invoices_path,
            failed_path,
            email_file_map,
            method,
            graph_client,
        )
        return True
    else:
        return False

def process_pdfs(
    pdf_files,
    invoices_path,
    excel_data,
    renamed_invoices_path,
    failed_path,
    email_file_map,
    graph_client: GraphEmailClient,
):
    for filename in pdf_files:
        file_path = os.path.join(invoices_path, filename)

        # First, try to find a match in the file name
        found_match, code = find_name_code_match(filename, excel_data)

        if found_match:
            handle_successful_match(
                filename,
                file_path,
                code,
                renamed_invoices_path,
                failed_path,
                email_file_map,
                'Filename',
                graph_client,
            )
            continue

        # If no match found in the file name, proceed with PyPDF2 extraction
        try:
            text = extract_text_pypdf2(file_path)
            found_match = process_pdf(
                filename,
                file_path,
                text,
                excel_data,
                renamed_invoices_path,
                failed_path,
                email_file_map,
                'PyPDF2',
                graph_client,
            )
        except Exception as e:
            found_match = False

        # If no match found using PyPDF2, proceed with OCR extraction
        if not found_match:
            text = extract_text_ocr(file_path)
            found_match = process_pdf(
                filename,
                file_path,
                text,
                excel_data,
                renamed_invoices_path,
                failed_path,
                email_file_map,
                'pytesseract',
                graph_client,
            )

        if not found_match:
            handle_failed_file(filename, file_path, failed_path, email_file_map, text, graph_client)

def read_csv_data(csv_file):
    with tempfile.NamedTemporaryFile(delete=False) as temp_file:
        shutil.copy2(csv_file, temp_file.name)
        try:
            csv_data = pd.read_csv(temp_file.name, encoding='utf-8', on_bad_lines='skip')
        except UnicodeDecodeError:
            csv_data = pd.read_csv(temp_file.name, encoding='ISO-8859-1', on_bad_lines='skip')
    os.unlink(temp_file.name)
    return csv_data

def pytesseract_main(updated_saved_attachments, email_file_map, graph_client: GraphEmailClient):
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
    for file_path, message_id in zip(updated_saved_attachments, email_file_map_copy.values()):
        email_file_map[os.path.basename(file_path)] = message_id

    pdf_files = [f for f in os.listdir(invoice_path) if f.lower().endswith('.pdf')]
    process_pdfs(
        pdf_files,
        invoice_path,
        csv_data,
        renamed_invoices_path,
        failed_path,
        email_file_map,
        graph_client,
    )


def main():
    graph_client = create_graph_email_client()

    # Load the excel data
    excel_data = pd.read_csv(
        r'C:\Users\Administrator\Better Bookkeeping Management\BBKM - Documents\BBKM Plan Management\Client Names.CSV',
        header=None,
    )

    # Define folder paths
    invoices_path = 'C:/BBKM_InvoiceSorter/Invoices'
    renamed_invoices_path = 'C:/BBKM_InvoiceSorter/Invoices/Renamed Invoices'
    failed_path = 'C:/BBKM_InvoiceSorter/Invoices/Failed'

    # Map of attachment filename to Graph message ID
    email_file_map = {}

    while True:
        # Get the PDF files from the invoices folder
        pdf_files = [f for f in os.listdir(invoices_path) if f.lower().endswith('.pdf')]

        # Process the PDF files
        process_pdfs(
            pdf_files,
            invoices_path,
            excel_data,
            renamed_invoices_path,
            failed_path,
            email_file_map,
            graph_client,
        )

        pytesseract_main(pdf_files, email_file_map, graph_client)

if __name__ == "__main__":
    main()
