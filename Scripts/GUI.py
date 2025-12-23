import queue
import threading
import tkinter as tk
from tkinter import ttk
from typing import List
import pythoncom
import pandas as pd
import os
import pyperclip
import sys
import datetime
from customtkinter import *
from tkinter import messagebox
from Main_Script import main

CSV_FILE_PATH = (
    "C:\\Users\\Administrator\\Better Bookkeeping Management\\BBKM - Documents\\"
    "BBKM Plan Management\\Client_Profiles.csv"
)
EXPECTED_COLUMNS = [
    "Client Code",
    "All Known Names",
    "NDIS Number",
    "Assigned Plan Manager",
]


def _normalize_client_profile_columns(df: pd.DataFrame) -> pd.DataFrame:
    trimmed_columns = [str(col).strip() for col in df.columns]
    df.columns = trimmed_columns

    rename_map = {}
    for idx, expected_name in enumerate(EXPECTED_COLUMNS):
        if expected_name not in df.columns and idx < len(df.columns):
            rename_map[df.columns[idx]] = expected_name
    if rename_map:
        df = df.rename(columns=rename_map)

    for col in EXPECTED_COLUMNS:
        if col not in df.columns:
            df[col] = ""

    return df[EXPECTED_COLUMNS]


def _split_known_names(raw_names: str) -> List[str]:
    separators = ["|", ",", ";"]
    names = [raw_names] if raw_names else []
    for sep in separators:
        names = [chunk for name in names for chunk in name.split(sep)]
    cleaned = []
    for name in names:
        trimmed = name.strip()
        if trimmed and trimmed not in cleaned:
            cleaned.append(trimmed)
    return cleaned


def _load_client_profiles() -> pd.DataFrame:
    if not os.path.isfile(CSV_FILE_PATH):
        return pd.DataFrame(columns=EXPECTED_COLUMNS)

    read_kwargs = {"dtype": str, "on_bad_lines": "skip", "na_filter": False}
    try:
        df = pd.read_csv(CSV_FILE_PATH, encoding="utf-8", **read_kwargs)
    except UnicodeDecodeError:
        df = pd.read_csv(CSV_FILE_PATH, encoding="ISO-8859-1", **read_kwargs)
    return _normalize_client_profile_columns(df)


def _save_client_profiles(df: pd.DataFrame) -> None:
    normalized = _normalize_client_profile_columns(df)
    normalized.to_csv(CSV_FILE_PATH, index=False)

class App:
    def __init__(self, root, stop_flag):
        self.root = root
        self.root.title("BBKM Invoice Sorter")
        self.root.geometry("520x900")

        self.stop_flag = stop_flag
        self._stop_requested = False

        # Initialize script_thread as None
        self.script_thread = None

        self.status_var = tk.StringVar(value="Status: Idle")

        self.main_button = CTkButton(root, text="Run Invoice Sorter", command=self.run_in_thread)
        self.main_button.pack(fill="x", padx=10, pady=10)

        self.stop_button = CTkButton(root, text="Stop Invoice Sorter", command=self.stop_main)
        self.stop_button.pack(fill="x", padx=10, pady=10)
        self.stop_button.configure(state="disabled")  # Disable the "Stop Invoice Sorter" button initially

        self.status_label = CTkLabel(root, textvariable=self.status_var)
        self.status_label.pack(fill="x", padx=10, pady=(0, 5))

        self.progress = ttk.Progressbar(root, mode="indeterminate")
        self.progress.pack(fill="x", padx=10)

        self.last_run_var = tk.StringVar(value="Last Run: Never")
        self.last_run_label = CTkLabel(root, textvariable=self.last_run_var)
        self.last_run_label.pack(fill="x", padx=10, pady=(5, 10))
        self.name_label = CTkLabel(root, text="Enter New Client Details Below")
        self.name_label.pack(padx=10, pady=5)

        self.name_placeholder = "Enter known names (use | to separate)"
        self.code_placeholder = "Enter client code"
        self.ndis_placeholder = "Enter NDIS number"
        self.plan_manager_placeholder = "Enter assigned plan manager"
        self.search_placeholder = "Search clients..."
        self.alias_placeholder = "Enter alias to add"

        self.name_entry = tk.Entry(root)
        self.name_entry.pack(fill="x", padx=10, pady=5)
        self.name_entry.insert(0, self.name_placeholder)
        self.name_entry.bind("<FocusIn>", self.clear_name_entry)
        self.name_entry.bind("<FocusOut>", self.restore_name_entry)

        self.code_entry = tk.Entry(root)
        self.code_entry.pack(fill="x", padx=10, pady=5)
        self.code_entry.insert(0, self.code_placeholder)
        self.code_entry.bind("<FocusIn>", self.clear_code_entry)
        self.code_entry.bind("<FocusOut>", self.restore_code_entry)

        self.ndis_entry = tk.Entry(root)
        self.ndis_entry.pack(fill="x", padx=10, pady=5)
        self.ndis_entry.insert(0, self.ndis_placeholder)
        self.ndis_entry.bind("<FocusIn>", self.clear_ndis_entry)
        self.ndis_entry.bind("<FocusOut>", self.restore_ndis_entry)

        self.plan_manager_entry = tk.Entry(root)
        self.plan_manager_entry.pack(fill="x", padx=10, pady=5)
        self.plan_manager_entry.insert(0, self.plan_manager_placeholder)
        self.plan_manager_entry.bind("<FocusIn>", self.clear_plan_manager_entry)
        self.plan_manager_entry.bind("<FocusOut>", self.restore_plan_manager_entry)

        self.add_button = CTkButton(root, text="Add to CSV", command=self.add_entry_to_csv)
        self.add_button.pack(fill="x", padx=10, pady=10)

        self.alias_entry = tk.Entry(root)
        self.alias_entry.pack(fill="x", padx=10, pady=(0, 5))
        self.alias_entry.insert(0, self.alias_placeholder)
        self.alias_entry.bind("<FocusIn>", self.clear_alias_entry)
        self.alias_entry.bind("<FocusOut>", self.restore_alias_entry)

        self.add_alias_button = CTkButton(
            root, text="Add Alias to Selected", command=self.add_alias_to_selected
        )
        self.add_alias_button.pack(fill="x", padx=10, pady=(0, 10))

        self.client_treeview = ttk.Treeview(
            root,
            columns=("Code", "Known Names", "NDIS Number", "Plan Manager"),
            show="headings",
            height=10,
        )
        self.client_treeview.heading("Code", text="Client Code")
        self.client_treeview.heading("Known Names", text="All Known Names")
        self.client_treeview.heading("NDIS Number", text="NDIS Number")
        self.client_treeview.heading("Plan Manager", text="Assigned Plan Manager")
        self.client_treeview.column("Code", width=100)
        self.client_treeview.column("Known Names", width=200)
        self.client_treeview.column("NDIS Number", width=120)
        self.client_treeview.column("Plan Manager", width=160)
        self.client_treeview.pack(fill="both", padx=10, pady=5)

        self.search_entry = tk.Entry(root)
        self.search_entry.pack(fill="x", padx=10, pady=5)
        self.search_entry.insert(0, self.search_placeholder)
        self.search_entry.bind("<FocusIn>", self.clear_search_entry)
        self.search_entry.bind("<FocusOut>", self.restore_search_entry)

        self.delete_button = CTkButton(root, text="Delete Entry", command=self.delete_entry_from_csv)
        self.delete_button.pack(fill="x", padx=10, pady=10)

        # Create a frame for the "Copy Name" and "Copy Code" buttons
        self.copy_frame = tk.Frame(root, bg="#1E1E1E")
        self.copy_frame.pack(fill="x", padx=10, pady=10)

        self.copy_name_button = CTkButton(self.copy_frame, text="Copy Name", command=self.copy_selected_name)
        self.copy_name_button.pack(side="left", expand=True)

        self.copy_code_button = CTkButton(self.copy_frame, text="Copy Code", command=self.copy_selected_code)
        self.copy_code_button.pack(side="left", expand=True)

        self.copy_both_button = CTkButton(root, text="Copy Both", command=self.copy_selected_both)
        self.copy_both_button.pack(fill="x", padx=10, pady=10)

        self.log_label = CTkLabel(root, text="Recent Activity")
        self.log_label.pack(fill="x", padx=10)

        self.log_frame = tk.Frame(root, bg="#1E1E1E")
        self.log_frame.pack(fill="both", expand=True, padx=10, pady=5)

        self.log_scrollbar = tk.Scrollbar(self.log_frame)
        self.log_scrollbar.pack(side="right", fill="y")

        self.log_text = tk.Text(
            self.log_frame,
            height=10,
            bg="#1E1E1E",
            fg="white",
            insertbackground="white",
            wrap="word",
            state="disabled",
            yscrollcommand=self.log_scrollbar.set,
        )
        self.log_text.pack(side="left", fill="both", expand=True)
        self.log_scrollbar.config(command=self.log_text.yview)

        self.open_log_button = CTkButton(root, text="Open Log File", command=self.open_log_file)
        self.open_log_button.pack(fill="x", padx=10, pady=(5, 10))

        # Bind the <KeyRelease> event to the search_clients method
        self.search_entry.bind("<KeyRelease>", self.search_clients)

        # Create a queue for communication
        self.queue = queue.Queue()

        # Populate the treeview when the GUI loads
        self.populate_client_listbox()

    def clear_name_entry(self, event):
        if self.name_entry.get() == self.name_placeholder:
            self.name_entry.delete(0, "end")

    def restore_name_entry(self, event):
        if self.name_entry.get() == "":
            self.name_entry.insert(0, self.name_placeholder)

    def clear_code_entry(self, event):
        if self.code_entry.get() == self.code_placeholder:
            self.code_entry.delete(0, "end")

    def restore_code_entry(self, event):
        if self.code_entry.get() == "":
            self.code_entry.insert(0, self.code_placeholder)

    def clear_ndis_entry(self, event):
        if self.ndis_entry.get() == self.ndis_placeholder:
            self.ndis_entry.delete(0, "end")

    def restore_ndis_entry(self, event):
        if self.ndis_entry.get() == "":
            self.ndis_entry.insert(0, self.ndis_placeholder)

    def clear_plan_manager_entry(self, event):
        if self.plan_manager_entry.get() == self.plan_manager_placeholder:
            self.plan_manager_entry.delete(0, "end")

    def restore_plan_manager_entry(self, event):
        if self.plan_manager_entry.get() == "":
            self.plan_manager_entry.insert(0, self.plan_manager_placeholder)

    def clear_alias_entry(self, event):
        if self.alias_entry.get() == self.alias_placeholder:
            self.alias_entry.delete(0, "end")

    def restore_alias_entry(self, event):
        if self.alias_entry.get() == "":
            self.alias_entry.insert(0, self.alias_placeholder)

    def clear_search_entry(self, event):
        if self.search_entry.get() == self.search_placeholder:
            self.search_entry.delete(0, "end")

    def restore_search_entry(self, event):
        if self.search_entry.get() == "":
            self.search_entry.delete(0, "end")
            self.search_entry.insert(0, self.search_placeholder)

    def add_entry_to_csv(self):
        raw_names = self.name_entry.get().strip()
        code = self.code_entry.get().strip()
        ndis_number = self.ndis_entry.get().strip()
        plan_manager = self.plan_manager_entry.get().strip()

        if raw_names == self.name_placeholder:
            raw_names = ""
        if code == self.code_placeholder:
            code = ""
        if ndis_number == self.ndis_placeholder:
            ndis_number = ""
        if plan_manager == self.plan_manager_placeholder:
            plan_manager = ""

        known_names = _split_known_names(raw_names)

        if not code or not known_names:
            messagebox.showerror(
                "Error",
                "Please enter a client code and at least one known name.",
            )
            return

        data_frame = _load_client_profiles()

        mask = data_frame["Client Code"].str.strip() == code
        if mask.any():
            current_row = data_frame[mask].iloc[0]
            existing_names = _split_known_names(current_row.get("All Known Names", ""))
            merged_names = _split_known_names(" | ".join(existing_names + known_names))
            combined_known_names = " | ".join(merged_names)
            updated_row = {
                "Client Code": code,
                "All Known Names": combined_known_names,
                "NDIS Number": ndis_number or str(current_row.get("NDIS Number", "")),
                "Assigned Plan Manager": plan_manager
                or str(current_row.get("Assigned Plan Manager", "")),
            }
            data_frame = data_frame[~mask]
            data_frame = pd.concat([data_frame, pd.DataFrame([updated_row])], axis=0)
        else:
            combined_known_names = " | ".join(known_names)
            new_entry = {
                "Client Code": code,
                "All Known Names": combined_known_names,
                "NDIS Number": ndis_number,
                "Assigned Plan Manager": plan_manager,
            }
            data_frame = pd.concat([data_frame, pd.DataFrame([new_entry])], axis=0)

        data_frame = data_frame.drop_duplicates(subset=["Client Code"], keep="last")
        _save_client_profiles(data_frame)

        self.name_entry.delete(0, "end")
        self.name_entry.insert(0, self.name_placeholder)
        self.code_entry.delete(0, "end")
        self.code_entry.insert(0, self.code_placeholder)
        self.ndis_entry.delete(0, "end")
        self.ndis_entry.insert(0, self.ndis_placeholder)
        self.plan_manager_entry.delete(0, "end")
        self.plan_manager_entry.insert(0, self.plan_manager_placeholder)

        self.populate_client_listbox()

    def populate_client_listbox(self):
        for i in self.client_treeview.get_children():
            self.client_treeview.delete(i)  # Clear the treeview

        # Read the CSV file and populate the treeview
        data_frame = _load_client_profiles()
        if not data_frame.empty:
            for _, row in data_frame.drop_duplicates(subset=["Client Code"], keep="last").iterrows():
                values = (
                    str(row.get("Client Code", "")),
                    str(row.get("All Known Names", "")),
                    str(row.get("NDIS Number", "")),
                    str(row.get("Assigned Plan Manager", "")),
                )
                self.client_treeview.insert("", "end", values=values)

    def search_clients(self, event=None):
        search_text = self.search_entry.get().lower()

        if search_text == self.search_placeholder.lower():
            search_text = ""

        # Clear the treeview
        for i in self.client_treeview.get_children():
            self.client_treeview.delete(i)

        # Search the client names and codes
        data_frame = _load_client_profiles()
        for _, row in data_frame.iterrows():
            values = [
                str(row.get("Client Code", "")),
                str(row.get("All Known Names", "")),
                str(row.get("NDIS Number", "")),
                str(row.get("Assigned Plan Manager", "")),
            ]
            if any(search_text in val.lower() for val in values):
                self.client_treeview.insert("", "end", values=values)

    def delete_entry_from_csv(self):
        cur_item = self.client_treeview.focus()
        if cur_item:
            selected_code = self.client_treeview.item(cur_item)['values'][0]

            data_frame = _load_client_profiles()
            mask = data_frame["Client Code"].str.strip() != str(selected_code)
            updated = data_frame[mask]

            _save_client_profiles(updated)

            self.name_entry.delete(0, 'end')
            self.code_entry.delete(0, 'end')
            self.ndis_entry.delete(0, 'end')
            self.plan_manager_entry.delete(0, 'end')
            self.alias_entry.delete(0, 'end')
            self.name_entry.insert(0, self.name_placeholder)
            self.code_entry.insert(0, self.code_placeholder)
            self.ndis_entry.insert(0, self.ndis_placeholder)
            self.plan_manager_entry.insert(0, self.plan_manager_placeholder)
            self.alias_entry.insert(0, self.alias_placeholder)

            # Repopulate the client listbox
            self.populate_client_listbox()

    def add_alias_to_selected(self):
        alias = self.alias_entry.get().strip()
        if alias == self.alias_placeholder:
            alias = ""

        if not alias:
            messagebox.showerror("Error", "Please enter an alias to add.")
            return

        cur_item = self.client_treeview.focus()
        if not cur_item:
            messagebox.showerror("Error", "Please select a client to add the alias to.")
            return

        selected_code = str(self.client_treeview.item(cur_item)["values"][0])

        data_frame = _load_client_profiles()
        mask = data_frame["Client Code"].str.strip() == selected_code
        if not mask.any():
            messagebox.showerror("Error", "Could not find the selected client in the CSV.")
            return

        existing_names = _split_known_names(data_frame[mask].iloc[0].get("All Known Names", ""))
        merged_names = _split_known_names(" | ".join(existing_names + [alias]))
        data_frame.loc[mask, "All Known Names"] = " | ".join(merged_names)

        _save_client_profiles(data_frame)

        self.alias_entry.delete(0, "end")
        self.alias_entry.insert(0, self.alias_placeholder)

        self.populate_client_listbox()

    def copy_selected_name(self):
        cur_item = self.client_treeview.focus()
        if cur_item:
            selected_names = self.client_treeview.item(cur_item)["values"][1]
            pyperclip.copy(selected_names)

    def copy_selected_code(self):
        cur_item = self.client_treeview.focus()
        if cur_item:
            selected_code = self.client_treeview.item(cur_item)["values"][0]
            pyperclip.copy(selected_code)

    def copy_selected_both(self):
        cur_item = self.client_treeview.focus()
        if cur_item:
            values = self.client_treeview.item(cur_item)["values"]
            selected_code = values[0]
            selected_names = values[1]
            parsed_names = _split_known_names(selected_names)
            primary_name = parsed_names[0] if parsed_names else selected_names
            selected_both = f"{selected_code} - {primary_name}"
            pyperclip.copy(selected_both)

    def check_script_finished(self):
        if not self.script_thread:
            return

        if not self.script_thread.is_alive():
            # The script thread has finished, do any post-processing here
            self.script_thread = None
            # Enable the "Run Invoice Sorter" button and disable the "Stop Invoice Sorter" button
            self.main_button.configure(state="normal")
            self.stop_button.configure(state="disabled")
            stop_requested = self._stop_requested or self.stop_flag.is_set()
            if stop_requested:
                self.status_var.set("Status: Stopped")
            else:
                self.status_var.set("Status: Idle")
            self.progress.stop()
            self.stop_flag.clear()
            self._stop_requested = False
        else:
            # The script is still running, schedule another check
            self.root.after(100, self.check_script_finished)

    def append_log(self, log_with_timestamp):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", log_with_timestamp + "\n")
        self.log_text.see("end")
        total_lines = int(self.log_text.index("end-1c").split(".")[0])
        if total_lines > 300:
            self.log_text.delete("1.0", f"{total_lines - 300}.0")
        self.log_text.configure(state="disabled")

    def update_log(self):
        # Check if there are logs in the queue
        while not self.queue.empty():
            log_item = self.queue.get()  # Get a log message from the queue
            if log_item is None:
                continue

            if isinstance(log_item, tuple):
                kind, payload = log_item
            else:
                kind, payload = "log", log_item

            if kind != "log":
                continue

            log_message = str(payload).strip()
            if not log_message:
                continue

            with open("log.txt", "a", encoding="utf-8", errors="replace") as file:
                file.write(log_message + "\n")

            self.append_log(log_message)

        self.root.after(500, self.update_log)

    def run_in_thread(self):
        if self.script_thread and self.script_thread.is_alive():
            return

        self.stop_flag.clear()
        self._stop_requested = False
        self.queue = queue.Queue()  # Create a queue to communicate with the main thread
        self.script_thread = threading.Thread(target=self.run_main, daemon=True)
        self.script_thread.start()
        self.main_button.configure(state="disabled")  # Disable the "Run Invoice Sorter" button
        self.stop_button.configure(state="normal")  # Enable the "Stop Invoice Sorter" button
        self.status_var.set("Status: Running")
        self.progress.start(10)
        self.last_run_var.set(
            "Last Run Started: "
            + datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        )
        self.root.after(100, self.check_script_finished)  # Start checking if the script has finished
        self.root.after(100, self.update_log)  # Start updating the log periodically

    def stop_main(self):
        if not self.script_thread or not self.script_thread.is_alive():
            return

        self.stop_button.configure(state="disabled")
        # Set the stop event flag to stop the script
        self.stop_flag.set()
        self._stop_requested = True
        sys.stdout = sys.__stdout__  # Assign the standard output to the original stream
        print("Stop command received...")
        self.status_var.set("Status: Stopped")


    def run_main(self):
        pythoncom.CoInitialize()  # Initialize the COM library
        print("Running main script...")
        try:
            main(self.stop_flag, self.queue)  # Pass the stop flag and log queue to the main script
        except Exception as e:
            print(f"Error occurred: {e}")
        finally:
            print("Main script finished")
            stop_requested = self.stop_flag.is_set()
            timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self.queue.put(None)  # Put a None value in the queue to signal the end of logs
            self.root.after(
                0,
                lambda stop_requested=stop_requested, timestamp=timestamp: self.finalize_run(
                    stop_requested, timestamp
                ),
            )

    def finalize_run(self, stop_requested, timestamp):
        if stop_requested:
            self.last_run_var.set(f"Last Run Stopped: {timestamp}")
        else:
            self.last_run_var.set(f"Last Run Completed: {timestamp}")

    def open_log_file(self):
        log_path = os.path.abspath("log.txt")
        if not os.path.exists(log_path):
            with open(log_path, "w"):
                pass

        try:
            os.startfile(log_path)
        except AttributeError:
            if sys.platform == "darwin":
                os.system(f"open '{log_path}'")
            else:
                os.system(f"xdg-open '{log_path}'")

set_appearance_mode("dark")
set_default_color_theme("dark-blue")

root = CTk()
stop_flag = threading.Event()  # Create an event flag for stopping the script
app = App(root, stop_flag)
root.mainloop()
