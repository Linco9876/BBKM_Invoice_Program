import os
import win32com.client
import shutil
import re
import filecmp
from win32com.client import Dispatch

save_path = "C:\\BBKM_InvoiceSorter\\Invoices"

def compare_files(file1, file2):
    if not os.path.exists(file1) or not os.path.exists(file2):
        return False
    return filecmp.cmp(file1, file2, shallow=False)

def save_attachments_from_outlook_folder(folder_name, save_path):
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    # Access shared mailbox 'accounts@bbkm.com.au'
    recipient = outlook.CreateRecipient("accounts@bbkm.com.au")
    if not recipient.Resolve():
        print("Unable to resolve shared mailbox: accounts@bbkm.com.au")
        return [], {}

    try:
        inbox = outlook.GetSharedDefaultFolder(recipient, 6)  # 6 = Inbox
    except Exception as e:
        print(f"Error accessing shared inbox: {e}")
        return [], {}

    saved_attachments = []
    email_file_map = {}

    items = inbox.Items
    items.Sort("[ReceivedTime]", False)  # Sort by oldest first

    for item in items:
        # Skip based on categories or flag
        if item.Categories in ["Attachment Extracted", "Skipped Email", "Reminder", "Service Agreement", "Quote", "Statement", "Credit Adj", "Remittance", "Doubled up"]:
            continue
        if item.FlagStatus == 2:
            continue

        subject = item.Subject or ""
        body = item.Body or ""
        sender_email = item.SenderEmailAddress.lower()

        # Keywords and classification
        if any(re.search(r'\bService Agreement\b', t, re.IGNORECASE) for t in [subject, body]):
            item.Categories = "Service Agreement"
            item.Save()
            print("Service Agreement Found")
            continue

        if re.search(r'\breminder\b', subject, re.IGNORECASE) or re.search(r'\breminder\b', body, re.IGNORECASE):
            item.Categories = "Reminder"
            item.Save()
            print("Reminder Found")
            continue

        if re.search(r'\bquote\b', subject, re.IGNORECASE) or re.search(r'\bquote\b', body, re.IGNORECASE):
            item.Categories = "Quote"
            item.Save()
            print("Quote Found")
            continue

        if re.search(r'\bOver-Due\b', subject, re.IGNORECASE) or re.search(r'\bOverdue\b', subject, re.IGNORECASE):
            item.Categories = "Reminder"
            item.Save()
            print("Reminder Found")
            continue

        if re.search(r'\bStatement\b', subject, re.IGNORECASE) and not re.search(r'\bActivity Statement\b', subject, re.IGNORECASE):
            item.Categories = "Statement"
            item.Save()
            print("Statement Found")
            continue

        if re.search(r'\bCredit Adj\b', subject, re.IGNORECASE) or re.search(r'\bCredit Adj\b', body, re.IGNORECASE):
            item.Categories = "Credit Adj"
            item.Save()
            print("Credit Adj Found")
            continue

        if "24 Pritchard Street" in subject or "24 Pritchard Street" in body:
            if "Activity Statement" not in subject:
                item.Categories = "Skipped Email"
                item.Save()
                print("Skipped Email Found")
                continue

        if any(domain in sender_email for domain in ["bbkm.com.au"]):
            if "Activity Statement" not in subject:
                item.Categories = "Skipped Email"
                item.Save()
                print("Skipped Domain")
                continue

        # Check for "remittance" in attachments
        skip_email_due_to_remittance = False
        for attachment in item.Attachments:
            if re.search(r'remittance', attachment.FileName, re.IGNORECASE):
                item.Categories = "Remittance"
                item.Save()
                print("Remittance Found")
                skip_email_due_to_remittance = True
                break
        if skip_email_due_to_remittance:
            continue

        # Process attachments
        attachment_saved = False
        has_attachments = False
        for attachment in item.Attachments:
            if attachment.FileName.lower().endswith(('.pdf', '.docx', '.doc')):
                # Identify business
                domain_map = {
                    "independenceaustralia.com": "independence australia",
                    "country-care.com.au": "country care",
                    "brightsky.com.au": "brightsky",
                    "visionaustralia.org": "visionaustralia",
                    "gsc.vic.gov.au": "Gannawarra Shire",
                    "alifesimplylived.com.au": "A Life Simply Lived",
                }
                business_name = next((name for domain, name in domain_map.items() if domain in sender_email), None)

                base_name, ext = os.path.splitext(attachment.FileName)
                file_name = f"{base_name} {business_name}{ext}" if business_name else attachment.FileName
                full_path = os.path.join(save_path, file_name)

                # Handle duplicates
                if os.path.exists(full_path):
                    temp_file_path = os.path.join(save_path, f"temp_{attachment.FileName}")
                    attachment.SaveAsFile(temp_file_path)
                    if compare_files(full_path, temp_file_path):
                        os.remove(temp_file_path)
                        item.Categories = "Doubled up"
                        item.Save()
                        continue
                    else:
                        new_path = os.path.join(save_path, f"new_{attachment.FileName}")
                        shutil.move(temp_file_path, new_path)
                        full_path = new_path

                try:
                    attachment.SaveAsFile(full_path)
                    attachment_saved = True
                    has_attachments = True
                    saved_attachments.append((item, full_path))
                    email_file_map[attachment.FileName] = item
                except Exception as e:
                    print(f"Error saving attachment: {e}")

        # Final categorisation
        if not has_attachments:
            item.FlagStatus = 2
            item.Save()
        elif attachment_saved:
            try:
                item.Categories = "Attachment Extracted"
                item.UnRead = False
                item.Save()
            except Exception as e:
                print(f"Error updating email: {e}")

        break  # Only process one email per run

    return saved_attachments, email_file_map
