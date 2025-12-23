import os
import sys
from docx2pdf import convert
from PIL import Image
import win32com.client
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
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            word.DisplayAlerts = False
            doc = word.Documents.Open(file_path)
            doc.SaveAs(output_pdf, FileFormat=17)  # 17 = PDF
            doc.Close()
            word.Quit()

        # --- Excel / CSV ---
        elif file_extension in ['.xlsx', '.xls', '.csv']:
            excel = win32com.client.Dispatch("Excel.Application")
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
            word = win32com.client.Dispatch("Word.Application")
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
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            msg = outlook.OpenSharedItem(file_path)
            msg.SaveAs(output_pdf, 17)  # 17 = PDF
            msg = None

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

        # Skip internal folders and the attachment manifest used by the
        # downloader. The manifest is not a document and should never be
        # processed as an invoice.
        if file_name in ["Failed", "Renamed Invoices"]:
            continue
        if filename.casefold() == "invoice_hashes.json":
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

def main():
    folder_path = r'C:\BBKM_InvoiceSorter\Invoices'
    saved_attachments = []
    saved_attachments = convert_files_to_pdf(folder_path, saved_attachments)

if __name__ == "__main__":
    main()
