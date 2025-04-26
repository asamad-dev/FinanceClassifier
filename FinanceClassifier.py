import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os

def process_excel(input_files, output_folder):
    for file_path in input_files:
        try:
            df = pd.read_excel(file_path)
            output_path = os.path.join(output_folder, f"processed_{os.path.basename(file_path)}")
            df.to_excel(output_path, index=False)
        except Exception as e:
            messagebox.showerror("Processing Error", f"Failed processing:\n{file_path}\n\nError:\n{str(e)}")

def select_files():
    files = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx *.xls")])
    if files:
        global selected_input_files
        selected_input_files = list(files)  # store real list
        input_files_display_var.set("; ".join(files))  # only for display

def select_output_folder():
    folder = filedialog.askdirectory()
    if folder:
        output_folder_var.set(folder)

def run_processing():
    if not selected_input_files or not output_folder_var.get():
        messagebox.showerror("Error", "Please select input files and output folder.")
        return
    process_excel(selected_input_files, output_folder_var.get())
    messagebox.showinfo("Success", "Processing completed!")

root = tk.Tk()
root.title("Excel Processor")

# Variables
selected_input_files = []  # real list of files
input_files_display_var = tk.StringVar()  # shown in entry
output_folder_var = tk.StringVar()

# Widgets
tk.Button(root, text="Select Excel Files", command=select_files).pack(pady=5)
tk.Entry(root, textvariable=input_files_display_var, width=80, state="readonly").pack(pady=5)

tk.Button(root, text="Select Output Folder", command=select_output_folder).pack(pady=5)
tk.Entry(root, textvariable=output_folder_var, width=80, state="readonly").pack(pady=5)

tk.Button(root, text="Run", command=run_processing).pack(pady=20)

root.mainloop()
