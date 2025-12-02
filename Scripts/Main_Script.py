import os
import time
import sys
import threading
from save_attachments_from_outlook_folder import (
    AuthConfigurationError,
    forward_emails_with_categories,
    save_attachments_from_outlook_folder,
)
from convert_files_to_pdf import convert_files_to_pdf
from pytesseractBBKMSCRIPT import pytesseract_main
from Move_to_OneDrive import move_files
import builtins
import datetime


class _QueueLogger:
    """File-like object that mirrors writes to a queue."""

    def __init__(self, file_obj, queue=None):
        self._file = file_obj
        self._queue = queue
        self._buffer = ""

    def write(self, message):
        if not message:
            return

        self._file.write(message)
        self._file.flush()

        if not self._queue:
            return

        self._buffer += message
        while "\n" in self._buffer:
            line, self._buffer = self._buffer.split("\n", 1)
            line = line.strip()
            if line:
                self._queue.put(("log", line))

    def flush(self):
        self._file.flush()
        if self._queue and self._buffer.strip():
            self._queue.put(("log", self._buffer.strip()))
        self._buffer = ""

    def isatty(self):
        return False

# Store the original print function
original_print = builtins.print

def print_with_timestamp(*args, **kwargs):
    timestamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    # Convert all arguments to string and concatenate with space separator
    all_args = ' '.join(str(arg) for arg in args)
    # Use original print here
    original_print(f"{timestamp} - {all_args}", **kwargs)

# Override the built-in print function
builtins.print = print_with_timestamp

def main(stop_flag, log_queue=None):
    log_file_path = "C:\\BBKM_InvoiceSorter\\BBKM_Logs\\output.log"  # Path to the log file

    invoices_path = "C:\\BBKM_InvoiceSorter\\Invoices"
    STA_invoices_path = "C:\\Users\\Administrator\\Better Bookkeeping Management\\BBKM - Documents\\BBKM Plan Management\\NDIS\\ZInvoices for lodgement\\Invoice Program\\STA Invoices"
    SRC_FOLDER = "C:\\BBKM_InvoiceSorter\\Invoices\\Renamed Invoices"
    DEST_FOLDER = "C:\\Users\\Administrator\\Better Bookkeeping Management\\BBKM - Documents\\BBKM Plan Management\\NDIS\\ZInvoices for lodgement\\Invoice Program"
    SRC_FOLDER_ATTEMPT = "C:\\Users\\Administrator\\Better Bookkeeping Management\\BBKM - Documents\\BBKM Plan Management\\NDIS\\ZInvoices for lodgement\\Invoice Program\\Attempt Code"
    DEST_FOLDER_ATTEMPT = "C:\\BBKM_InvoiceSorter\\Invoices"
    SRC_FOLDER_FAILED = "C:\\BBKM_InvoiceSorter\\Invoices\\Failed"
    DEST_FOLDER_FAILED = "C:\\Users\\Administrator\\Better Bookkeeping Management\\BBKM - Documents\\BBKM Plan Management\\NDIS\\ZInvoices for lodgement\\Invoice Program\\Failed to Code"

    max_retries = 3
    retry_delay = 5  # in seconds

    original_stdout = sys.stdout

    # Open the log file in append mode
    with open(log_file_path, "a", encoding="utf-8", errors="replace") as log_file:
        # Redirect the standard output to a tee logger
        logger = _QueueLogger(log_file, log_queue)
        sys.stdout = logger

        try:
            while not stop_flag.is_set():
                try:
                    # Run script 1 to save attachments from Outlook
                    saved_attachments, email_file_map = save_attachments_from_outlook_folder("invoices", invoices_path)

                    # Run script 2 move emails with specific categories to info@bbkm.com.au inbox
                    forward_emails_with_categories(
                        "info@bbkm.com.au",
                        [
                            "Service Agreement",
                            "Reminder",
                            "quote",
                            "Remittance",
                            "Statement",
                            "Caution Email",
                            "Credit Adj",
                        ],
                    )

                    # Run script 3 to convert files to PDF
                    updated_saved_attachments = convert_files_to_pdf(invoices_path, saved_attachments)

                    # Run script 4 to extract text from PDFs and rename them based on client names
                    pytesseract_main(updated_saved_attachments, email_file_map)

                    # Run script 5 to move files to the appropriate subfolders in OneDrive
                    move_files(SRC_FOLDER_ATTEMPT, DEST_FOLDER_ATTEMPT)
                    move_files(SRC_FOLDER_FAILED, DEST_FOLDER_FAILED)
                    move_files(SRC_FOLDER, DEST_FOLDER)

                    # Pause for 5 Seconds
                    time.sleep(5)

                except AuthConfigurationError as auth_error:
                    print(f"Fatal authentication error: {auth_error}")
                    print("Halting processing until credentials are updated.")
                    break
                except Exception as e:
                    print(f"Error occurred: {e}")
                    max_retries -= 1

                    if max_retries <= 0:
                        print("Max retries reached. Continuing...")
                        # Reset max_retries to its original value for the next loop
                        max_retries = 5

                    print(f"Retrying in {retry_delay} seconds...")
                    time.sleep(retry_delay)
        finally:
            logger.flush()
            sys.stdout = original_stdout

if __name__ == "__main__":
    stop_flag = threading.Event()  # Create an event flag for stopping the script
    main(stop_flag)
