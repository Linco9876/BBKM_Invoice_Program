import queue
import threading
import tkinter as tk
from tkinter import ttk
import pythoncom
import pandas as pd
import csv
import os
import pyperclip
import sys
import datetime
from customtkinter import *
from tkinter import messagebox
from Main_Script import main

CSV_FILE_PATH = (
    "C:\\Users\\Administrator\\Better Bookkeeping Management\\BBKM - Documents\\"
    "BBKM Plan Management\\Client Names.csv"
)

class App:
    def __init__(self, root, stop_flag):
        self.root = root
        self.root.title("BBKM Invoice Sorter")
        self.root.geometry("360x780")

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
        self.name_label = CTkLabel(root, text="Enter New Names and Codes Below")
        self.name_label.pack(padx=10, pady=5)

        self.name_entry = tk.Entry(root)
        self.name_entry.pack(fill="x", padx=10, pady=5)
        self.name_entry.insert(0, "Enter name...")
        self.name_entry.bind("<FocusIn>", self.clear_name_entry)
        self.name_entry.bind("<FocusOut>", self.restore_name_entry)

        self.code_entry = tk.Entry(root)
        self.code_entry.pack(fill="x", padx=10, pady=5)
        self.code_entry.insert(0, "Enter code...")
        self.code_entry.bind("<FocusIn>", self.clear_code_entry)
        self.code_entry.bind("<FocusOut>", self.restore_code_entry)

        self.add_button = CTkButton(root, text="Add to CSV", command=self.add_entry_to_csv)
        self.add_button.pack(fill="x", padx=10, pady=10)

        self.client_treeview = ttk.Treeview(root, columns=("Name", "Code"), show='headings', height=8)
        self.client_treeview.heading("Name", text="Name")
        self.client_treeview.heading("Code", text="Code")
        self.client_treeview.pack(fill="both", padx=10, pady=5)

        self.search_entry = tk.Entry(root)
        self.search_entry.pack(fill="x", padx=10, pady=5)
        self.search_entry.insert(0, "Search Names...")
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
        if self.name_entry.get() == "Enter name...":
            self.name_entry.delete(0, "end")

    def restore_name_entry(self, event):
        if self.name_entry.get() == "":
            self.name_entry.insert(0, "Enter name...")

    def clear_code_entry(self, event):
        if self.code_entry.get() == "Enter code...":
            self.code_entry.delete(0, "end")

    def restore_code_entry(self, event):
        if self.code_entry.get() == "":
            self.code_entry.insert(0, "Enter code...")

    def clear_search_entry(self, event):
        if self.search_entry.get() == "Search Names...":
            self.search_entry.delete(0, "end")

    def restore_search_entry(self, event):
        if self.search_entry.get() == "":
            self.search_entry.delete(0, "end")
            self.search_entry.insert(0, "Search Names...")

    def add_entry_to_csv(self):
        name = self.name_entry.get()
        code = self.code_entry.get()

        if name and code:
            new_entry = {'A': [name], 'B': [code]}  # Use default column names 'A' and 'B'
            new_data_frame = pd.DataFrame(new_entry)

            # Read the CSV file
            data_frame = pd.read_csv(CSV_FILE_PATH, header=None, names=['A', 'B'])

            # Concatenate the new entry to the existing dataframe
            updated_data_frame = pd.concat([data_frame, new_data_frame], axis=0)

            # Drop duplicate rows
            updated_data_frame = updated_data_frame.drop_duplicates()

            # Save the updated DataFrame back to the CSV file
            updated_data_frame.to_csv(CSV_FILE_PATH, header=False, index=False)

            # Clear the entry boxes
            self.name_entry.delete(0, 'end')
            self.code_entry.delete(0, 'end')

            # Repopulate the client listbox
            self.populate_client_listbox()
        else:
            messagebox.showerror("Error", "Please enter both a name and a code.")

    def populate_client_listbox(self):
        for i in self.client_treeview.get_children():
            self.client_treeview.delete(i)  # Clear the treeview

        # Read the CSV file and populate the treeview
        if os.path.isfile(CSV_FILE_PATH):
            data_frame = pd.read_csv(CSV_FILE_PATH, header=None, names=['A', 'B'])
            unique_entries = data_frame.drop_duplicates().values.tolist()
            
            for entry in unique_entries:
                name, code = entry
                self.client_treeview.insert('', 'end', values=(name, code))

    def search_clients(self, event=None):
        search_text = self.search_entry.get().lower()

        # Clear the treeview
        for i in self.client_treeview.get_children():
            self.client_treeview.delete(i)

        # Search the client names and codes
        if os.path.isfile(CSV_FILE_PATH):
            data_frame = pd.read_csv(CSV_FILE_PATH, header=None, names=['A', 'B'])
            for _, row in data_frame.iterrows():
                name = row['A']
                code = row['B']
                if search_text in name.lower() or search_text in code.lower():
                    self.client_treeview.insert('', 'end', values=(name, code))

    def delete_entry_from_csv(self):
        cur_item = self.client_treeview.focus()
        if cur_item:
            selected_name, selected_code = self.client_treeview.item(cur_item)['values']

            # Read the CSV file
            data_frame = pd.read_csv(CSV_FILE_PATH, header=None, names=['A', 'B'])

            # Get index of rows with the selected name and code
            index_names = data_frame[(data_frame['A'] == selected_name) & (data_frame['B'] == selected_code)].index

            # Delete these row indexes from dataFrame
            data_frame.drop(index_names, inplace=True)

            # Save the updated DataFrame back to the CSV file
            data_frame.to_csv(CSV_FILE_PATH, header=False, index=False)

            # Clear the entry boxes
            self.name_entry.delete(0, 'end')
            self.code_entry.delete(0, 'end')

            # Repopulate the client listbox
            self.populate_client_listbox()

    def copy_selected_name(self):
        cur_item = self.client_treeview.focus()
        if cur_item:
            selected_name = self.client_treeview.item(cur_item)["values"][0]
            pyperclip.copy(selected_name)

    def copy_selected_code(self):
        cur_item = self.client_treeview.focus()
        if cur_item:
            selected_code = self.client_treeview.item(cur_item)["values"][1]
            pyperclip.copy(selected_code)

    def copy_selected_both(self):
        cur_item = self.client_treeview.focus()
        if cur_item:
            selected_name, selected_code = self.client_treeview.item(cur_item)["values"]
            selected_both = f"{selected_code} - {selected_name}"
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
