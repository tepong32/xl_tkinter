import tkinter as tk
from tkinter import ttk, messagebox
import ttkbootstrap as tb
from ttkbootstrap.constants import *
from openpyxl import load_workbook, Workbook
import os

FILE_NAME = "data.xlsx"

def ensure_excel_file():
    """Ensure Excel file exists with headers."""
    if not os.path.exists(FILE_NAME):
        wb = Workbook()
        ws = wb.active
        ws.title = "Data"
        ws.append(["Name", "Age", "Email"])
        wb.save(FILE_NAME)

def load_excel_data(tree):
    """Load Excel data into Treeview (latest entries first)."""
    tree.delete(*tree.get_children())
    wb = load_workbook(FILE_NAME)
    sheet = wb.active
    rows = list(sheet.iter_rows(values_only=True))
    for row in reversed(rows):
        if any(row):  # skip empty rows
            tree.insert('', tk.END, values=row)
    wb.close()

def insert_data(name, age, email, tree):
    """Insert new data into Excel and refresh the tree."""
    if not name or not age or not email:
        messagebox.showerror("Error", "All fields are required.")
        return
    try:
        wb = load_workbook(FILE_NAME)
        ws = wb.active
        ws.append([name, age, email])
        wb.save(FILE_NAME)
        wb.close()
        load_excel_data(tree)
        messagebox.showinfo("Success", "Data added successfully!")
    except Exception as e:
        messagebox.showerror("Error", str(e))

def center_window(win, width, height):
    """Center the window on the screen."""
    screen_width = win.winfo_screenwidth()
    screen_height = win.winfo_screenheight()
    x = (screen_width // 2) - (width // 2)
    y = (screen_height // 2) - (height // 2)
    win.geometry(f"{width}x{height}+{x}+{y}")

def main():
    ensure_excel_file()

    root = tb.Window(themename="flatly")
    root.title("Data Entry App")
    center_window(root, 1280, 960)
    root.minsize(1280, 960)

    frame = ttk.Frame(root, padding=10)
    frame.pack(fill=BOTH, expand=True)

    # Input Fields
    form_frame = ttk.LabelFrame(frame, text="Add New Record", padding=10)
    form_frame.pack(fill=X, pady=10)

    ttk.Label(form_frame, text="Name:").grid(row=0, column=0, padx=5, pady=5, sticky=E)
    ttk.Label(form_frame, text="Age:").grid(row=1, column=0, padx=5, pady=5, sticky=E)
    ttk.Label(form_frame, text="Email:").grid(row=2, column=0, padx=5, pady=5, sticky=E)

    name_var = tk.StringVar()
    age_var = tk.StringVar()
    email_var = tk.StringVar()

    name_entry = ttk.Entry(form_frame, textvariable=name_var, width=30)
    age_entry = ttk.Entry(form_frame, textvariable=age_var, width=30)
    email_entry = ttk.Entry(form_frame, textvariable=email_var, width=30)

    name_entry.grid(row=0, column=1, padx=5, pady=5)
    age_entry.grid(row=1, column=1, padx=5, pady=5)
    email_entry.grid(row=2, column=1, padx=5, pady=5)

    tree_frame = ttk.LabelFrame(frame, text="Records", padding=10)
    tree_frame.pack(fill=BOTH, expand=True, pady=10)

    columns = ("Name", "Age", "Email")
    tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=20)
    for col in columns:
        tree.heading(col, text=col)
        tree.column(col, width=200, anchor=CENTER)

    scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
    tree.configure(yscroll=scrollbar.set)
    tree.pack(side=LEFT, fill=BOTH, expand=True)
    scrollbar.pack(side=RIGHT, fill=Y)

    load_excel_data(tree)

    add_button = tb.Button(form_frame, text="Add Record", bootstyle=SUCCESS,
                           command=lambda: insert_data(name_var.get(), age_var.get(), email_var.get(), tree))
    add_button.grid(row=3, column=0, columnspan=2, pady=10)

    root.mainloop()

if __name__ == "__main__":
    main()
