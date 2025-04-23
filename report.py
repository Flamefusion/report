import tkinter as tk
from tkinter import filedialog, messagebox
import ttkbootstrap as ttk
import os
from report_module import generate_report  # Assuming you have this module already

def run_gui():
    def select_file():
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if path:
            file_path_var.set(path)

    def process():
        file_path = file_path_var.get()
        report_for = report_type_var.get()
        if not file_path or not os.path.exists(file_path):
            messagebox.showerror("Error", "Please select a valid Excel file.")
            return
        output = generate_report(file_path, report_for)
        messagebox.showinfo("Success", f"Report generated:\n{output}")

    # Create GUI with ttkbootstrap style
    root = ttk.Window(themename="darkly")  # Using darkly theme for a modern look
    root.title("Manufacturing Report Generator")
    root.geometry("500x300")

    file_path_var = tk.StringVar()
    report_type_var = tk.StringVar(value="3DE TECH")

    # Add Widgets with modernized design
    ttk.Label(root, text="Select Excel File:").pack(pady=10, padx=20)
    ttk.Entry(root, textvariable=file_path_var, width=40).pack(pady=5)
    ttk.Button(root, text="Browse", command=select_file).pack(pady=10)

    ttk.Label(root, text="Select Report Type:").pack(pady=10)
    ttk.Combobox(root, textvariable=report_type_var, values=["3DE TECH", "IHC"], state="readonly").pack(pady=5)

    ttk.Button(root, text="Generate Report", command=process).pack(pady=15)

    root.mainloop()

if __name__ == '__main__':
    run_gui()
