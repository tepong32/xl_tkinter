import tkinter as tk
from tkinter import ttk, messagebox
import ttkbootstrap as tb
from ttkbootstrap.constants import *
import openpyxl
import os

class DataEntryApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Data Entry App")
        self.root.minsize(1280, 960)  # 🔧 Set minimum desktop-like size
        self.style = tb.Style("darkly")
        
        self.excel_path = "./people.xlsx"
        self.ensure_excel_file()

        # Configure main frame for responsive layout
        self.frame = ttk.Frame(root, padding=10)
        self.frame.pack(fill="both", expand=True)
        self.frame.columnconfigure(1, weight=1)
        self.frame.rowconfigure(0, weight=1)

        # === Left panel: Data entry ===
        self.widgets_frame = ttk.LabelFrame(self.frame, text="Insert Row", padding=10)
        self.widgets_frame.grid(row=0, column=0, padx=10, pady=10, sticky="n")

        self.name_entry = ttk.Entry(self.widgets_frame)
        self.name_entry.insert(0, "Name")
        self.name_entry.bind("<FocusIn>", lambda e: self.clear_placeholder(self.name_entry, "Name"))
        self.name_entry.grid(row=0, column=0, padx=5, pady=5, sticky="ew")

        self.age_spinbox = ttk.Spinbox(self.widgets_frame, from_=18, to=100)
        self.age_spinbox.insert(0, "Age")
        self.age_spinbox.grid(row=1, column=0, padx=5, pady=5, sticky="ew")

        combo_list = ["Subscribed", "Not Subscribed", "Other"]
        self.status_combobox = ttk.Combobox(self.widgets_frame, values=combo_list)
        self.status_combobox.current(0)
        self.status_combobox.grid(row=2, column=0, padx=5, pady=5, sticky="ew")

        self.cb_var = tk.BooleanVar()
        self.checkbutton = ttk.Checkbutton(self.widgets_frame, text="Employed", variable=self.cb_var)
        self.checkbutton.grid(row=3, column=0, padx=5, pady=5, sticky="w")

        insert_btn = ttk.Button(self.widgets_frame, text="Insert", bootstyle=SUCCESS, command=self.insert_row)
        insert_btn.grid(row=4, column=0, padx=5, pady=10, sticky="ew")

        ttk.Separator(self.widgets_frame).grid(row=5, column=0, padx=5, pady=10, sticky="ew")

        self.mode_var = tk.BooleanVar()
        mode_switch = ttk.Checkbutton(
            self.widgets_frame, text="Toggle Theme", variable=self.mode_var,
            bootstyle="round-toggle", command=self.toggle_mode)
        mode_switch.grid(row=6, column=0, padx=5, pady=10, sticky="ew")

        # === Right panel: TreeView ===
        self.tree_frame = ttk.Frame(self.frame)
        self.tree_frame.grid(row=0, column=1, padx=10, pady=10, sticky="nsew")

        self.tree_scroll = ttk.Scrollbar(self.tree_frame)
        self.tree_scroll.pack(side="right", fill="y")

        self.cols = ("Name", "Age", "Subscription", "Employment")
        self.tree = ttk.Treeview(
            self.tree_frame, show="headings", yscrollcommand=self.tree_scroll.set,
            columns=self.cols, height=15)
        self.tree_scroll.config(command=self.tree.yview)

        for col in self.cols:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=120, anchor="center")

        self.tree.pack(fill="both", expand=True)
        self.tree.bind("<<TreeviewSelect>>", self.on_select)

        self.load_data()

    def clear_placeholder(self, entry, placeholder):
        if entry.get() == placeholder:
            entry.delete(0, "end")

    def ensure_excel_file(self):
        if not os.path.exists(self.excel_path):
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.append(["Name", "Age", "Subscription", "Employment"])
            wb.save(self.excel_path)

    def load_data(self):
        self.tree.delete(*self.tree.get_children())
        wb = openpyxl.load_workbook(self.excel_path)
        ws = wb.active
        rows = list(ws.iter_rows(min_row=2, values_only=True))
        # 🔁 Reverse so newest (last rows) appear first
        for row in reversed(rows):
            self.tree.insert("", tk.END, values=row)
        wb.close()

    def insert_row(self):
        name = self.name_entry.get().strip()
        age_value = self.age_spinbox.get()
        sub_status = self.status_combobox.get()
        emp_status = "Employed" if self.cb_var.get() else "Unemployed"

        if not name or name == "Name":
            messagebox.showerror("Error", "Please enter a valid name.")
            return
        try:
            age = int(age_value)
        except ValueError:
            messagebox.showerror("Error", "Age must be a number.")
            return

        wb = openpyxl.load_workbook(self.excel_path)
        ws = wb.active
        ws.append([name, age, sub_status, emp_status])
        wb.save(self.excel_path)
        wb.close()

        # 🆕 Insert new row at the top instead of the bottom
        self.tree.insert("", 0, values=(name, age, sub_status, emp_status))

        self.name_entry.delete(0, "end")
        self.name_entry.insert(0, "Name")
        self.age_spinbox.delete(0, "end")
        self.age_spinbox.insert(0, "Age")
        self.status_combobox.current(0)
        self.cb_var.set(False)

    def on_select(self, event):
        selected_item = self.tree.selection()
        if selected_item:
            values = self.tree.item(selected_item, "values")
            print("Selected:", values)

    def toggle_mode(self):
        theme = "flatly" if self.mode_var.get() else "darkly"
        self.style.theme_use(theme)


if __name__ == "__main__":
    root = tb.Window(themename="darkly")
    app = DataEntryApp(root)
    root.mainloop()
