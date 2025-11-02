import win32com.client

def forward_emails_with_categories(to_address, categories):
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

        # Shared mailbox access
        accounts_folder = outlook.Folders.Item("accounts@bbkm.com.au")
        info_folder = outlook.Folders.Item(to_address)

        accounts_inbox = accounts_folder.Folders["Inbox"]
        complete_invoices = accounts_folder.Folders["Complete Invoices"]
        info_inbox = info_folder.Folders["Inbox"]

        messages = accounts_inbox.Items
        messages.Sort("[ReceivedTime]", True)  # Sort newest first

        for message in messages:
            if message.Class != 43:  # 43 = MailItem
                continue

            if not message.Categories:
                continue

            # Check if any category matches (case-insensitive)
            msg_categories = [c.strip().lower() for c in message.Categories.split(",")]
            if any(cat.lower() in msg_categories for cat in categories):
                # Copy first, then move
                message_copy = message.Copy()
                message.UnRead = False
                message.Move(complete_invoices)
                print(f"Moved: '{message.Subject}' to '{complete_invoices.Name}'.")

                message_copy.Move(info_inbox)
                print(f"Copied: '{message.Subject}' to '{info_inbox.Name}'.")

    except Exception as e:
        print(f"Error occurred: {e}")

if __name__ == "__main__":
    forward_emails_with_categories(
        to_address="info@bbkm.com.au",
        categories=[
            "Service Agreement",
            "Reminder",
            "Quote",
            "Remittance",
            "Statement",
            "Caution Email",
            "Credit Adj"
        ]
    )
