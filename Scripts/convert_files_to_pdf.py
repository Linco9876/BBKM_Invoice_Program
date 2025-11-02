import os
import re
import textwrap
from email import policy
from email.parser import BytesParser
from html import unescape

from docx2pdf import convert
from PIL import Image
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import shutil
import pandas as pd

# --- Supported extensions ---
SUPPORTED_EXTENSIONS = [
    '.pdf', '.docx', '.doc', '.xlsx', '.xls', '.csv',
    '.png', '.jpg', '.jpeg', '.tif', '.tiff', '.bmp', '.heic',
    '.txt', '.rtf', '.odt', '.ods', '.html', '.eml'
]

def convert_to_pdf(file_path, output_pdf):
    file_name, file_extension = os.path.splitext(file_path)
    file_extension = file_extension.lower()

    try:
        # --- Word formats ---
        if file_extension == '.docx':
            convert(file_path, output_pdf)
        elif file_extension == '.doc':
            word = _get_win32_dispatch()("Word.Application")
            word.Visible = False
            word.DisplayAlerts = False
            doc = word.Documents.Open(file_path)
            doc.SaveAs(output_pdf, FileFormat=17)  # 17 = PDF
            doc.Close()
            word.Quit()

        # --- Excel / CSV ---
        elif file_extension in ['.xlsx', '.xls', '.csv']:
            excel = _get_win32_dispatch()("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            wb = excel.Workbooks.Open(file_path)
            ws_index_list = [1]
            wb.WorkSheets(ws_index_list).Select()
            wb.ActiveSheet.ExportAsFixedFormat(0, output_pdf)
            wb.Close()
            excel.Quit()

        # --- Text formats ---
        elif file_extension in ['.txt', '.rtf', '.odt', '.html']:
            word = _get_win32_dispatch()("Word.Application")
            word.Visible = False
            word.DisplayAlerts = False
            doc = word.Documents.Open(file_path)
            doc.SaveAs(output_pdf, FileFormat=17)
            doc.Close()
            word.Quit()

        # --- Images (raster formats) ---
        elif file_extension in ['.png', '.jpg', '.jpeg', '.tif', '.tiff', '.bmp', '.heic']:
            image = Image.open(file_path)
            image.save(output_pdf, "PDF", resolution=100.0)

        # --- EML (Outlook email message) ---
        elif file_extension == '.eml':
            convert_eml_to_pdf(file_path, output_pdf)

        # --- Already PDF ---
        elif file_extension == '.pdf':
            shutil.copy(file_path, output_pdf)

        else:
            raise ValueError(f"Unsupported file format: {file_extension}")

    except Exception as e:
        raise RuntimeError(f"Error converting {file_name}{file_extension}: {e}")

def convert_files_to_pdf(folder_path, saved_attachments):
    updated_saved_attachments = []
    failed_folder_path = os.path.join(folder_path, "Failed")

    os.makedirs(failed_folder_path, exist_ok=True)

    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        file_name, file_extension = os.path.splitext(filename)
        file_extension = file_extension.lower()

        # Skip internal folders
        if file_name in ["Failed", "Renamed Invoices"]:
            continue

        output_pdf = os.path.join(folder_path, file_name + '.pdf')

        if file_extension == '.pdf':
            updated_saved_attachments.append(file_path)
            continue

        if file_extension in SUPPORTED_EXTENSIONS:
            try:
                convert_to_pdf(file_path, output_pdf)
                os.remove(file_path)
                updated_saved_attachments.append(output_pdf)
                print(f"Converted: {filename} → {os.path.basename(output_pdf)}")
            except Exception as e:
                print(f"Error converting {filename}: {e}")
                try:
                    shutil.move(file_path, os.path.join(failed_folder_path, filename))
                except Exception as e2:
                    print(f"Failed to move {filename} to 'Failed': {e2}")
        else:
            print(f"Unsupported file format: {filename}")
            try:
                shutil.move(file_path, os.path.join(failed_folder_path, filename))
            except Exception as e:
                print(f"Failed to move unsupported {filename}: {e}")

    return updated_saved_attachments

def _get_win32_dispatch():
    try:
        from win32com.client import Dispatch  # type: ignore
    except ImportError as exc:
        raise RuntimeError(
            "The pywin32 package is required to convert Word, Excel, and text-based files."
        ) from exc
    return Dispatch

def convert_eml_to_pdf(file_path, output_pdf):
    with open(file_path, 'rb') as eml_file:
        message = BytesParser(policy=policy.default).parse(eml_file)

    body_text = _extract_eml_body(message)
    header_lines = _format_eml_headers(message)

    _render_text_to_pdf(header_lines + [""] + body_text.splitlines(), output_pdf)

def _extract_eml_body(message):
    text_parts = []
    html_parts = []

    for part in message.walk():
        if part.get_content_maintype() == 'multipart':
            continue

        payload = part.get_payload(decode=True)
        if payload is None:
            continue

        charset = part.get_content_charset() or 'utf-8'
        try:
            decoded_payload = payload.decode(charset, errors='replace')
        except LookupError:
            decoded_payload = payload.decode('utf-8', errors='replace')

        content_type = part.get_content_type()
        if content_type == 'text/plain':
            text_parts.append(decoded_payload)
        elif content_type == 'text/html':
            html_parts.append(decoded_payload)

    if text_parts:
        return "\n\n".join(text_parts).strip()

    for html in html_parts:
        text = _html_to_text(html)
        if text:
            return text

    payload = message.get_payload(decode=True)
    if payload:
        charset = message.get_content_charset() or 'utf-8'
        try:
            return payload.decode(charset, errors='replace').strip()
        except LookupError:
            return payload.decode('utf-8', errors='replace').strip()

    return "(No content)"

def _html_to_text(html_content):
    text = re.sub(r'<(br|p)[^>]*?>', '\n', html_content, flags=re.IGNORECASE)
    text = re.sub(r'<style.*?>.*?</style>', '', text, flags=re.DOTALL | re.IGNORECASE)
    text = re.sub(r'<script.*?>.*?</script>', '', text, flags=re.DOTALL | re.IGNORECASE)
    text = re.sub(r'<[^>]+>', '', text)
    text = unescape(text)
    return text.strip()

def _format_eml_headers(message):
    headers = []
    for label in ['Subject', 'From', 'To', 'Date']:
        value = message.get(label, '')
        if value:
            headers.append(f"{label}: {value}")
    return headers

def _render_text_to_pdf(lines, output_pdf):
    pdf = canvas.Canvas(output_pdf, pagesize=letter)
    width, height = letter
    x_margin = 72
    y_margin = 72
    max_width = width - (2 * x_margin)
    line_height = 14

    y_position = height - y_margin

    for line in lines:
        wrapped_lines = textwrap.wrap(line, width=int(max_width / 7)) or ['']
        for wrapped_line in wrapped_lines:
            if y_position < y_margin:
                pdf.showPage()
                y_position = height - y_margin
            pdf.drawString(x_margin, y_position, wrapped_line)
            y_position -= line_height

    pdf.save()

def main():
    folder_path = r'C:\BBKM_InvoiceSorter\Invoices'
    saved_attachments = []
    saved_attachments = convert_files_to_pdf(folder_path, saved_attachments)

if __name__ == "__main__":
    main()
