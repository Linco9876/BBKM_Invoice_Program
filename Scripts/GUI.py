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

CSV_FILE_PATH = "C:\\Users\\Administrator\\Better Bookkeeping Management\\BBKM - Documents\\BBKM Plan Management\\Client Names.csv"

class App:
    def __init__(self, root, stop_flag):
        self.root = root
        self.root.title("BBKM Invoice Sorter")
        self.root.geometry("310x650")

        self.stop_flag = stop_flag

        # Initialize script_thread as None
        self.script_thread = None

        self.main_button = CTkButton(root, text="Run Invoice Sorter", command=self.run_in_thread)
        self.main_button.pack(fill="x", padx=10, pady=10)

        self.stop_button = CTkButton(root, text="Stop Invoice Sorter", command=self.stop_main)
        self.stop_button.pack(fill="x", padx=10, pady=10)
        self.stop_button.configure(state="disabled")  # Disable the "Stop Invoice Sorter" button initially
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

        self.client_treeview = ttk.Treeview(root, columns=("Name", "Code"), show='headings')
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

        # Bind the <KeyRelease> event to the search_clients method
        self.search_entry.bind("<KeyRelease>", self.search_clients)

        # Create a queue for communication
        self.queue = queue.Queue()

    def copy_selected_name(self, event):
        selected_index = self.client_listbox.curselection()
        if selected_index:
            selected_name = self.client_listbox.get(selected_index)
            pyperclip.copy(selected_name)

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
            selected_name = self.client_treeview.item(cur_item)['values'][0]
            pyperclip.copy(selected_name)

    def copy_selected_code(self):
        cur_item = self.client_treeview.focus()
        if cur_item:
            selected_code = self.client_treeview.item(cur_item)['values'][1]
            pyperclip.copy(selected_code)

    def copy_selected_both(self):
        cur_item = self.client_treeview.focus()
        if cur_item:
            selected_code, selected_name = self.client_treeview.item(cur_item)['values']
            selected_both = f"{selected_code} - {selected_name}"
            pyperclip.copy(selected_both)

    def check_script_finished(self):
        if self.script_thread and not self.script_thread.is_alive():
            # The script thread has finished, do any post-processing here
            self.script_thread = None
            # Enable the "Run Invoice Sorter" button and disable the "Stop Invoice Sorter" button
            self.main_button.configure(state="normal")
            self.stop_button.configure(state="disabled")
        else:
            # The script is still running, schedule another check
            self.root.after(100, self.check_script_finished)

    def update_log(self):
        # Check if there are logs in the queue
        while not self.queue.empty():
            log = self.queue.get()  # Get a log message from the queue
            if log:
                timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")  # Get current timestamp
                log_with_timestamp = f"{timestamp} - {log}"  # Add timestamp to log message

                # Print the log with timestamp to the console
                print(log_with_timestamp)

                # Write the log with timestamp to the log file
                with open("log.txt", "a") as file:
                    file.write(f"[{timestamp}] {log}\n")

    def run_in_thread(self):
        self.queue = queue.Queue()  # Create a queue to communicate with the main thread
        self.script_thread = threading.Thread(target=self.run_main)
        self.script_thread.start()
        self.main_button.configure(state="disabled")  # Disable the "Run Invoice Sorter" button
        self.stop_button.configure(state="normal")  # Enable the "Stop Invoice Sorter" button
        self.root.after(100, self.check_script_finished)  # Start checking if the script has finished
        self.root.after(100, self.update_log)  # Start updating the log periodically

    def stop_main(self):
        self.stop_button.configure(state="disabled")
        # Set the stop event flag to stop the script
        self.stop_flag.set()
        sys.stdout = sys.__stdout__  # Assign the standard output to the original stream
        print("Stop command received...")
        if self.script_thread:
            self.script_thread.join()  # Wait for the script thread to finish
            self.script_thread = None  # Reset the script thread
        self.main_button.configure(state="normal")

    def run_main(self):
        pythoncom.CoInitialize()  # Initialize the COM library
        print("Running main script...")
        # Reset flag at the start of the script
        self.should_stop = False
        try:
             main(self.stop_flag)  # Pass the stop flag to the main script
        except Exception as e:
            print(f"Error occurred: {e}")
        finally:
            print("Main script finished")
            self.queue.put(None)  # Put a None value in the queue to signal the end of logs


set_appearance_mode("dark")
set_default_color_theme("dark-blue")

root = CTk()
stop_flag = threading.Event()  # Create an event flag for stopping the script
app = App(root, stop_flag)
root.mainloop()