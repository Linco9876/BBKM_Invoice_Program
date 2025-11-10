from save_attachments_from_outlook_folder import forward_emails_with_categories


__all__ = ["forward_emails_with_categories"]


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
