import os
import re
from collections import Counter, defaultdict
from datetime import date, datetime

import customtkinter as ctk
from tkcalendar import DateEntry
import matplotlib
import matplotlib.pyplot as plt
import pandas as pd
from matplotlib.backends.backend_pdf import PdfPages

matplotlib.use("Agg")

LOG_FILE_PATH = "C:/BBKM_InvoiceSorter/BBKM_Logs/output.log"
OUTPUT_DIR = "C:/BBKM_InvoiceSorter/BBKM_Logs/Reports"

CATEGORY_PATTERNS = {
    "Service Agreement": re.compile(r"\bService Agreement Found\b", re.IGNORECASE),
    "Reminder": re.compile(r"\bReminder Found\b", re.IGNORECASE),
    "Remittance": re.compile(r"\bRemittance Found\b", re.IGNORECASE),
    "Statement": re.compile(r"\bStatement Found\b", re.IGNORECASE),
    "Credit Adj": re.compile(r"\bCredit Adj Found\b", re.IGNORECASE),
}

SORTING_RULES = {
    "Streamline Invoices": re.compile(r"^Moved to Streamline\b", re.IGNORECASE),
    "Manual Lodgement": re.compile(r"^Moved to Manual Lodgement\b", re.IGNORECASE),
    "AT&Consumables": re.compile(r"^Moved to AT&Consumables\b", re.IGNORECASE),
    "STA and Assistance": re.compile(r"^Moved to STA and Assistance\b", re.IGNORECASE),
    "New Provider": re.compile(r"^Moved to New Provider\b", re.IGNORECASE),
    "Custom Vendor": re.compile(r"^Moved to custom vendor\b", re.IGNORECASE),
    "Failed to Code": re.compile(r"^Moved failed invoice to|^Moved unknown failed invoice\b", re.IGNORECASE),
    "NDIS Activity Statement": re.compile(r"^NDIS statement moved\b", re.IGNORECASE),
}

PLAN_MANAGER_PATTERN = re.compile(r"\bunder (?P<manager>[^:]+):", re.IGNORECASE)
LOG_LINE_PATTERN = re.compile(r"^(?P<timestamp>\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}) - (?P<message>.*)$")


def _parse_log_lines(log_path: str, start: date, end: date):
    if not os.path.exists(log_path):
        return []

    matches = []
    with open(log_path, encoding="utf-8", errors="replace") as log_file:
        for line in log_file:
            line = line.strip()
            match = LOG_LINE_PATTERN.match(line)
            if not match:
                continue
            timestamp = datetime.strptime(match.group("timestamp"), "%Y-%m-%d %H:%M:%S")
            if start <= timestamp.date() <= end:
                matches.append(match.group("message"))
    return matches


def _summarize_messages(messages):
    category_counts = Counter()
    sorting_counts = Counter()
    plan_manager_counts = Counter()

    for message in messages:
        for label, pattern in CATEGORY_PATTERNS.items():
            if pattern.search(message):
                category_counts[label] += 1

        for label, pattern in SORTING_RULES.items():
            if pattern.search(message):
                sorting_counts[label] += 1
                manager_match = PLAN_MANAGER_PATTERN.search(message)
                if manager_match:
                    plan_manager = manager_match.group("manager").strip() or "Unassigned"
                    plan_manager_counts[plan_manager] += 1
                break

    processed_total = sum(sorting_counts.values())

    summary = {
        "Total Processed": processed_total,
        "Service Agreements": category_counts.get("Service Agreement", 0),
        "Reminders": category_counts.get("Reminder", 0),
        "Remittances": category_counts.get("Remittance", 0),
        "Statements": category_counts.get("Statement", 0),
        "Credit Adj": category_counts.get("Credit Adj", 0),
    }

    return summary, category_counts, sorting_counts, plan_manager_counts


def _table_figure(df: pd.DataFrame, title: str):
    fig, ax = plt.subplots(figsize=(11, 7.5))
    ax.axis("off")
    ax.set_title(title, fontsize=14, pad=18)

    table = ax.table(
        cellText=df.values,
        colLabels=df.columns,
        rowLabels=df.index,
        cellLoc="center",
        loc="center",
    )
    table.auto_set_font_size(False)
    table.set_fontsize(10)
    table.scale(1.0, 1.4)
    return fig


def generate_report(start_date: date, end_date: date, log_path: str = LOG_FILE_PATH):
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    messages = _parse_log_lines(log_path, start_date, end_date)
    summary, category_counts, sorting_counts, plan_manager_counts = _summarize_messages(messages)

    date_label = f"{start_date}_to_{end_date}"
    pdf_path = os.path.join(OUTPUT_DIR, f"Invoice_Report_{date_label}.pdf")

    summary_df = pd.DataFrame.from_dict(summary, orient="index", columns=["Count"])
    category_df = pd.DataFrame.from_dict(category_counts, orient="index", columns=["Count"]).sort_index()
    sorting_df = pd.DataFrame.from_dict(sorting_counts, orient="index", columns=["Count"]).sort_values(
        "Count",
        ascending=False,
    )
    plan_manager_df = pd.DataFrame.from_dict(plan_manager_counts, orient="index", columns=["Count"]).sort_values(
        "Count",
        ascending=False,
    )

    summary_df.to_csv(os.path.join(OUTPUT_DIR, f"Summary_{date_label}.csv"))
    category_df.to_csv(os.path.join(OUTPUT_DIR, f"Email_Categories_{date_label}.csv"))
    sorting_df.to_csv(os.path.join(OUTPUT_DIR, f"Sorting_Breakdown_{date_label}.csv"))
    plan_manager_df.to_csv(os.path.join(OUTPUT_DIR, f"Plan_Manager_Breakdown_{date_label}.csv"))

    with PdfPages(pdf_path) as pdf:
        pdf.savefig(_table_figure(summary_df, "Invoice Processing Summary"))
        pdf.savefig(_table_figure(category_df, "Email Category Counts"))
        pdf.savefig(_table_figure(sorting_df, "Invoice Sorting Breakdown"))
        pdf.savefig(_table_figure(plan_manager_df, "Invoices by Plan Manager"))

    return pdf_path


def create_gui():
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("dark-blue")

    root = ctk.CTk()
    root.title("Invoice Report Generator")
    root.geometry("420x320")

    padding = {"padx": 12, "pady": 6}

    ctk.CTkLabel(root, text="Start Date").grid(row=0, column=0, sticky="w", **padding)
    start_cal = DateEntry(
        root,
        width=18,
        background="#1a1a1a",
        foreground="white",
        borderwidth=1,
        date_pattern="yyyy-mm-dd",
    )
    start_cal.grid(row=0, column=1, **padding)

    ctk.CTkLabel(root, text="End Date").grid(row=1, column=0, sticky="w", **padding)
    end_cal = DateEntry(
        root,
        width=18,
        background="#1a1a1a",
        foreground="white",
        borderwidth=1,
        date_pattern="yyyy-mm-dd",
    )
    end_cal.grid(row=1, column=1, **padding)

    status_var = ctk.StringVar(value="Select a date range and generate a report.")
    status_label = ctk.CTkLabel(root, textvariable=status_var, wraplength=360, justify="left")
    status_label.grid(row=2, column=0, columnspan=2, sticky="w", **padding)

    def run_report():
        start_date = datetime.strptime(start_cal.get(), "%Y-%m-%d").date()
        end_date = datetime.strptime(end_cal.get(), "%Y-%m-%d").date()
        if start_date > end_date:
            status_var.set("Start date must be before the end date.")
            return

        report_path = generate_report(start_date, end_date)
        status_var.set(f"Report created: {report_path}")
        if os.path.exists(report_path):
            os.startfile(report_path)

    def open_output():
        if os.path.exists(OUTPUT_DIR):
            os.startfile(OUTPUT_DIR)

    ctk.CTkButton(root, text="Generate Report", command=run_report).grid(
        row=3,
        column=0,
        columnspan=2,
        sticky="ew",
        padx=12,
        pady=12,
    )
    ctk.CTkButton(root, text="Open Reports Folder", command=open_output).grid(
        row=4,
        column=0,
        columnspan=2,
        sticky="ew",
        padx=12,
        pady=(0, 12),
    )

    root.grid_columnconfigure(0, weight=1)
    root.grid_columnconfigure(1, weight=1)

    root.mainloop()


if __name__ == "__main__":
    create_gui()
