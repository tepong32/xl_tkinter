# test.py
# tEppy's Data Entry App — rebuilt, precision-safe, rounded-value markers
# Fixed: Restored missing _update_status method to prevent crash on startup.

import json
import os
import re
import sys
import warnings
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from ttkbootstrap import Window, Style
from openpyxl import load_workbook, Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime, date
import pandas as pd
from decimal import Decimal, ROUND_HALF_UP

APP_TITLE = "tEppy's Data Entry (Excel Companion with validation)"

# -------------------------
# Utilities
# -------------------------
def resource_path(relative_path):
    """Get absolute path to resource (works for dev and PyInstaller)."""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# -------------------------
# Tooltip Helper
# -------------------------
class ToolTip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tip_window = None
        widget.bind("<Enter>", self.show_tip)
        widget.bind("<Leave>", self.hide_tip)

    def show_tip(self, event=None):
        if self.tip_window or not self.text:
            return
        try:
            x, y, cx, cy = self.widget.bbox("insert")
        except Exception:
            x, y = 0, 0
        x += self.widget.winfo_rootx() + 20
        y += self.widget.winfo_rooty() + 20
        self.tip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        label = ttk.Label(tw, text=self.text, background="#ffffe0", padding=(6,3), relief="solid")
        label.pack()

    def hide_tip(self, event=None):
        if self.tip_window:
            self.tip_window.destroy()
            self.tip_window = None

# -------------------------
# Excel loader (universal)
# -------------------------
def load_any_excel(path: str, app_instance=None):
    ext = os.path.splitext(path)[1].lower()
    try:
        if ext in [".xlsx", ".xlsm"]:
            with warnings.catch_warnings(record=True) as w:
                warnings.simplefilter("always")
                wb = load_workbook(path, data_only=True)

            # Check for embedded images which might break saving
            unsupported_formats = (".wmf", ".emf", ".tiff", ".bmp")
            for ws in wb.worksheets:
                for image in getattr(ws, "_images", []):
                    if any(fmt in str(getattr(image, 'path', '')).lower() for fmt in unsupported_formats):
                        if app_instance:
                            app_instance._show_temp_warning(
                                "⚠️ Workbook contains embedded images (WMF/EMF).", 5000
                            )
                        break
            return wb

        elif ext == ".xlsb":
            df = pd.read_excel(path, engine="pyxlsb")
        elif ext == ".xls":
            df = pd.read_excel(path, engine="xlrd")
        elif ext == ".ods":
            df = pd.read_excel(path, engine="odf")
        else:
            raise ValueError(f"Unsupported file format: {ext}")

        # Convert dataframe back to openpyxl workbook
        wb = Workbook()
        ws = wb.active
        for r in dataframe_to_rows(df, index=False, header=True):
            ws.append(r)
        return wb

    except Exception as e:
        raise RuntimeError(f"Failed to load file ({ext}): {e}")

# -------------------------
# Validation helpers
# -------------------------
EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

def try_parse_date(value: str):
    if value is None:
        return None
    s = str(value).strip()
    if s == "":
        return None
    formats = ["%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y", "%Y/%m/%d"]
    for fmt in formats:
        try:
            return datetime.strptime(s, fmt).date()
        except Exception:
            continue
    try:
        return datetime.fromisoformat(s).date()
    except Exception:
        return None

def is_numeric(value: str):
    if value is None:
        return False
    s = str(value).strip()
    if s == "":
        return False
    try:
        float(s.replace(",", ""))
        return True
    except Exception:
        return False

def normalize_numeric(value: str, fmt: str = "integer"):
    """
    Parse numeric input. Returns int or float.
    """
    if value is None or str(value).strip() == "":
        return None
    s = str(value).replace(",", "").strip()
    try:
        num = float(s)
    except Exception:
        raise ValueError(f"Invalid numeric input: {value}")
    
    if fmt == "decimal":
        return num
    else:
        # If integer format is requested, we try to cast to int if it's safe
        if num.is_integer():
            return int(num)
        return num 

# -------------------------
# Precision & display helpers
# -------------------------
def format_value_for_display(value, rule=None, decimal_places=2):
    if value is None:
        return ""
    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, float):
        if rule and rule.get("format") == "decimal":
            return f"{value:.{decimal_places}f}"
        else:
            if value.is_integer():
                return str(int(value))
            return str(value)
    if isinstance(value, int):
        return str(value)
    return str(value)

def round_half_up(value, decimal_places=2):
    """
    Excel-style rounding (0-4 down, 5-9 up).
    """
    try:
        q = Decimal("1." + "0" * decimal_places)
        return float(Decimal(str(value)).quantize(q, rounding=ROUND_HALF_UP))
    except:
        return value

def detect_precision_mismatch(raw_text: str, parsed_value, decimal_places=2):
    """
    Returns True if the raw input differs from the rounded version.
    """
    if raw_text is None:
        return False
    s = str(raw_text).strip()
    if s == "":
        return False
    try:
        raw_num = float(s.replace(",", ""))
    except Exception:
        return False
    
    rounded = round_half_up(raw_num, decimal_places)
    return abs(raw_num - rounded) > 1e-9

def has_excess_precision(raw_text: str, decimal_limit=2) -> bool:
    """
    Checks if a string representation of a number has more than 'decimal_limit' digits.
    """
    if not raw_text: 
        return False
    s = str(raw_text).replace(",", "").strip()
    if "." in s:
        decimals = s.split(".")[1]
        if len(decimals) > decimal_limit:
            return True
    return False

# -------------------------
# Main App
# -------------------------
class DynamicExcelApp:
    def __init__(self, root: Window):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("1100x720")

        # State
        self.workbook = None
        self.filepath = None
        self.active_sheet_name = None
        self.current_sheet = None
        self.headers = []
        self.input_entries = []
        self.unsaved_changes = False
        self.original_editing_values = {}
        self.mode = "add"

        # Validation rules and helpers
        self.validation_rules = []
        self.current_rules = {}
        self.input_order = []

        # Shadow store: keys are (sheet_name, display_row_index, col_index)
        self.shadow_values = {}

        # Generated Icon for rounded values (Red Triangle)
        self._generate_rounded_icon()

        # UI Construction
        self.style = ttk.Style()
        self._create_menu()
        self._create_toolbar()
        self._create_top_frame()
        self._create_bottom_frame()
        self._create_statusbar()

        # Bind events & shortcuts
        self._bind_events()
        self._bind_shortcuts()

        # Prompt at start
        self._prompt_open_file_on_startup()

    def _generate_rounded_icon(self):
        """Creates a simple in-memory icon for the Treeview using tkinter commands."""
        try:
            self.rounded_icon = tk.PhotoImage(width=12, height=12)
            # Draw a simple red triangle
            data = []
            for y in range(12):
                row = []
                for x in range(12):
                    if 1 <= x <= 10 and 1 <= y <= 10 and y >= abs(x - 6) + 2:
                         row.append("#d9534f") # Bootstrap danger color
                    else:
                        row.append("None") # Transparent
                data.append("{" + " ".join(row) + "}")
            self.rounded_icon.put(" ".join(data))
        except Exception:
            self.rounded_icon = None # Fallback

    # -------------------------
    # Shadow helpers
    # -------------------------
    def _shadow_set(self, sheet_name, display_row_index, col_index, raw_text, parsed_value, rounded_flag=False):
        key = (sheet_name, int(display_row_index), int(col_index))
        self.shadow_values[key] = {
            "raw": None if raw_text is None else str(raw_text),
            "value": parsed_value,
            "rounded_flag": bool(rounded_flag)
        }

    def _shadow_get(self, sheet_name, display_row_index, col_index):
        key = (sheet_name, int(display_row_index), int(col_index))
        return self.shadow_values.get(key)

    # -------------------------
    # Tooltip Logic
    # -------------------------
    def _show_tooltip(self, widget, text):
        if hasattr(self, "_tooltip"):
            self._tooltip.destroy()

        x, y, _, _ = widget.bbox("insert")
        x += widget.winfo_rootx() + 20
        y += widget.winfo_rooty() + 20

        self._tooltip = tk.Toplevel(widget)
        self._tooltip.wm_overrideredirect(True)
        self._tooltip.attributes("-topmost", True)

        label = tk.Label(
            self._tooltip,
            text=text,
            background="#ffffe0",
            relief="solid",
            borderwidth=1,
            font=("Segoe UI", 9)
        )
        label.pack(ipadx=6, ipady=3)

        self._tooltip.geometry(f"+{x}+{y}")

    def _hide_tooltip(self):
        if hasattr(self, "_tooltip"):
            self._tooltip.destroy()
            del self._tooltip

    # -------------------------
    # Event bindings
    # -------------------------
    def _bind_events(self):
        self.tree.bind("<Double-1>", self.on_tree_double_click)
        self.tree.bind("<Motion>", self._on_tree_hover)
        self.tree.bind("<<TreeviewSelect>>", self._on_row_select)
        self.tree.bind("<Delete>", lambda e: self.delete_selected_row())

    def _on_row_select(self, event=None):
        try:
            selection = self.tree.selection()
            if not selection: return
            item_id = selection[0]
            self.selected_item = item_id
            row_index = int(self.tree.index(item_id)) + 1
            self._set_status(f"Selected row {row_index}")
        except Exception:
            pass

    def edit_selected_row(self):
        sel = getattr(self, "selected_item", None) or self.tree.focus()
        if sel:
            self.tree.selection_set(sel)
            self.on_tree_double_click(None)

    def new_file(self):
        wb = Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        self.workbook = wb
        self.filepath = None
        self.active_sheet_name = ws.title
        self._populate_sheet_selector()
        self._load_active_sheet()
        self._update_status("New workbook created.", "success")

    def _bind_shortcuts(self):
        self.root.bind("<Control-o>", lambda e: self.open_file())
        self.root.bind("<Control-s>", lambda e: self.save_file())
        self.root.bind("<Control-S>", lambda e: self.save_file_as())
        self.root.bind("<Control-n>", lambda e: self.new_file())
        self.root.bind("<F2>", lambda e: self.edit_selected_row())
        self.root.bind("<Escape>", lambda e: self.reset_to_add_mode())

    # -------------------------
    # UI Construction Methods
    # -------------------------
    def _create_menu(self):
        menubar = tk.Menu(self.root)
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="New", command=self.new_file)
        file_menu.add_command(label="Open...", command=self.open_file)
        file_menu.add_command(label="Save", command=self.save_file)
        file_menu.add_command(label="Save As...", command=self.save_file_as)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.on_close)
        menubar.add_cascade(label="File", menu=file_menu)

        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="Help & Instructions", command=self._show_help)
        help_menu.add_separator()
        help_menu.add_command(label="About", command=self._show_about)
        menubar.add_cascade(label="Help", menu=help_menu)

        self.root.config(menu=menubar)

    def _create_toolbar(self):
        toolbar = ttk.Frame(self.root, padding=(8,6))
        toolbar.pack(side=tk.TOP, fill=tk.X)

        delete_btn = ttk.Button(toolbar, text="🗑 Delete Row", command=self.delete_selected_row, style="danger.TButton")
        delete_btn.pack(side=tk.LEFT, padx=(0,10))

        self.add_button = ttk.Button(toolbar, text="➕ Add Row", command=self.add_row_from_inputs, style="success.TButton")
        self.add_button.pack(side=tk.LEFT, padx=(0,14))

        ttk.Label(toolbar, text="Sheet:", bootstyle="secondary").pack(side=tk.LEFT, padx=(8,4))
        self.sheet_combo = ttk.Combobox(toolbar, state="readonly", width=28)
        self.sheet_combo.pack(side=tk.LEFT, padx=(0,8))
        self.sheet_combo.bind("<<ComboboxSelected>>", self.on_sheet_change)

        spacer = ttk.Label(toolbar, text="")
        spacer.pack(side=tk.LEFT, expand=True)

        right_grp = ttk.Frame(toolbar, padding=(8,4))
        right_grp.pack(side=tk.RIGHT)
        right_grp.config(relief=tk.GROOVE, borderwidth=1)

        self.auto_save_var = tk.BooleanVar(value=False)
        auto_save_chk = ttk.Checkbutton(right_grp, text="Auto-Save", variable=self.auto_save_var, style="primary.TCheckbutton")
        auto_save_chk.pack(side=tk.LEFT, padx=(0,8))

        ttk.Separator(right_grp, orient="vertical").pack(side=tk.LEFT, fill=tk.Y, padx=(4,10))

        ttk.Label(right_grp, text="Theme:", bootstyle="secondary").pack(side=tk.LEFT, padx=(4,4))
        self.theme_combo = ttk.Combobox(right_grp, values=Style().theme_names(), state="readonly", width=15)
        self.theme_combo.set(Style().theme_use())
        self.theme_combo.bind("<<ComboboxSelected>>", self.on_theme_change)
        self.theme_combo.pack(side=tk.LEFT, padx=(0,6))

        help_btn = ttk.Button(right_grp, text="❓ Help", command=self._show_help, style="info.TButton", width=8)
        help_btn.pack(side=tk.LEFT, padx=(10,0))
        ToolTip(help_btn, "View usage instructions")
        self._add_hover_effect(help_btn)

    def _create_statusbar(self):
        self.status_var = tk.StringVar(value="Ready")
        self.status_label = ttk.Label(self.root, textvariable=self.status_var, anchor="e", padding=(6,2), bootstyle="secondary")
        self.status_label.pack(side=tk.BOTTOM, fill=tk.X)

    def _set_status(self, message):
        """Simple status update (used by hover events)"""
        if not message:
            self.status_var.set("Ready")
            try:
                self.status_label.configure(background="#f8f9fa", foreground="#333333")
            except Exception:
                pass
            return
        try:
            self.status_var.set(message)
            self.status_label.configure(background="#fff3cd", foreground="#664d03")
        except Exception:
            self.status_var.set(message)

    def _update_status(self, message, level="info", duration=5000):
        """Main status update with color coding"""
        colors = {
            "info": "#f8f9fa",
            "success": "#d1e7dd",
            "warning": "#fff3cd",
            "error": "#f8d7da",
        }
        fg_colors = {
            "info": "#333333",
            "success": "#0f5132",
            "warning": "#664d03",
            "error": "#842029",
        }
        try:
            self.status_var.set(message)
            bg = colors.get(level, colors["info"])
            fg = fg_colors.get(level, fg_colors["info"])
            self.status_label.configure(background=bg, foreground=fg)
            if duration > 0:
                self.root.after(duration, lambda: self._fade_status())
        except Exception:
            pass

    def _fade_status(self):
        """Resets status to neutral state"""
        try:
            self.status_label.configure(background="#f8f9fa", foreground="#333333")
            self.status_var.set("Ready")
        except: pass

    def _create_top_frame(self):
        self.top_frame = ttk.Frame(self.root)
        self.top_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=(6,0))

        self.input_canvas = tk.Canvas(self.top_frame, height=140)
        self.input_canvas.pack(side=tk.TOP, fill=tk.X, expand=True)
        self.input_scrollbar = ttk.Scrollbar(self.top_frame, orient="horizontal", command=self.input_canvas.xview)
        self.input_scrollbar.pack(side=tk.TOP, fill=tk.X)
        self.input_canvas.configure(xscrollcommand=self.input_scrollbar.set)

        self.inputs_inner = ttk.Frame(self.input_canvas)
        self.input_canvas.create_window((0,0), window=self.inputs_inner, anchor="nw")
        self.inputs_inner.bind("<Configure>", lambda e: self.input_canvas.configure(scrollregion=self.input_canvas.bbox("all")))

    def _create_bottom_frame(self):
        bottom_frame = ttk.Frame(self.root)
        bottom_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.filter_frame = ttk.Frame(bottom_frame)
        self.filter_frame.pack(side=tk.TOP, fill=tk.X, pady=(0,3))

        tree_frame = ttk.Frame(bottom_frame)
        tree_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        self.tree = ttk.Treeview(tree_frame, show="headings")
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        vsb.pack(side=tk.LEFT, fill=tk.Y)
        self.tree.configure(yscrollcommand=vsb.set)

        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.configure(xscrollcommand=hsb.set)

        self.all_rows = []
        self.filter_entries = []
        self.column_ids = []

        self.tree.bind("<Configure>", lambda e: self._adjust_filter_widths())
        self.tree.bind("<ButtonRelease-1>", lambda e: self._adjust_filter_widths())

    # -------------------------
    # Hover: Show raw vs rounded
    # -------------------------
    def _on_tree_hover(self, event):
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            self._set_status("")
            return

        row_id = self.tree.identify_row(event.y)
        col_id = self.tree.identify_column(event.x)
        if not row_id or not col_id:
            self._set_status("")
            return
        try:
            col_index = int(col_id.replace("#",""))
            row_index_excel = int(self.tree.index(row_id)) + 1
        except Exception:
            self._set_status("")
            return

        shadow = self._shadow_get(self.current_sheet, row_index_excel, col_index)
        if shadow and shadow.get("rounded_flag"):
            raw = shadow.get("raw")
            displayed = self.tree.set(row_id, col_index - 1)
            msg = f"Rounded-off value: original '{raw}' → displayed '{displayed}'"
            self._set_status(msg)
            try:
                self._show_tooltip(self.tree, f"Original: {raw}\nDisplayed: {displayed}")
            except:
                pass
        else:
            self._set_status("")
            self._hide_tooltip()

    # -------------------------
    # Filters
    # -------------------------
    def _create_filter_row(self):
        for w in self.filter_frame.winfo_children():
            w.destroy()
        self.filter_entries.clear()
        self.column_ids = list(self.tree["columns"])
        if not self.column_ids: return
        for col_id in self.column_ids:
            entry = ttk.Entry(self.filter_frame)
            entry.pack(side=tk.LEFT, padx=1, fill=tk.X, expand=True)
            entry.insert(0, "")
            entry.bind("<KeyRelease>", lambda e: self._apply_filters())
            self.filter_entries.append(entry)
        self.root.after(100, self._adjust_filter_widths)

    def _adjust_filter_widths(self):
        if not self.filter_entries or not self.column_ids: return
        for i, col_id in enumerate(self.column_ids):
            try:
                width = int(self.tree.column(col_id, "width"))
            except Exception:
                width = 100
            self.filter_entries[i].config(width=max(8, width // 10))

    def _apply_filters(self):
        if not self.all_rows: return
        filters = [f.get().strip().lower() for f in self.filter_entries]
        if all(f == "" for f in filters):
            self._reload_tree_from_cache()
            return
        filtered = []
        for row in self.all_rows:
            if all((f in str(row[i]).lower() if f else True) for i,f in enumerate(filters)):
                filtered.append(row)
        self._reload_tree_from_cache(filtered)

    def _reload_tree_from_cache(self, rows=None):
        self.tree.delete(*self.tree.get_children())
        display_rows = rows if rows is not None else self.all_rows
        for row in display_rows:
            row_extended = list(row) + [""] * (len(self.headers) - len(row))
            self.tree.insert("", tk.END, values=row_extended)

    def _clear_treeview(self):
        for col in self.tree["columns"]:
            self.tree.heading(col, text="")
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = ()

    # -------------------------
    # File Operations
    # -------------------------
    def _prompt_open_file_on_startup(self):
        answer = messagebox.askyesno("Open file", "Is your template ready for loading?")
        if answer:
            self.open_file()
        else:
            self.status_var.set("Ready. Use File -> Open or Ctrl+O to open a spreadsheet.")

    def open_file(self):
        filetypes = [
            ("All supported files", "*.xlsx *.xlsm *.xlsb *.xls *.ods"),
            ("Excel files", "*.xlsx;*.xlsm;*.xlsb;*.xls"),
        ]
        path = filedialog.askopenfilename(title="Open spreadsheet", filetypes=filetypes)
        if not path: return
        try:
            wb = load_any_excel(path, app_instance=self)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open file:\n{e}")
            return

        self.workbook = wb
        self.filepath = path
        try:
            self.active_sheet_name = self.workbook.active.title
        except Exception:
            self.active_sheet_name = self.workbook.sheetnames[0] if self.workbook.sheetnames else None
        self._populate_sheet_selector()
        self._load_active_sheet()
        self.unsaved_changes = False
        
        ext = os.path.splitext(path)[1].lower()
        if ext not in [".xlsx", ".xlsm"]:
            self._update_status(f"Opened {os.path.basename(path)} (converted to in-memory .xlsx)")
            messagebox.showinfo("Format Notice", "This file was opened from a non-.xlsx format.\nIt will be saved as .xlsx when you save changes.")
        else:
            self._update_status(f"Opened: {os.path.basename(path)}")

    def _populate_sheet_selector(self):
        if not self.workbook:
            self.sheet_combo["values"] = []
            return
        names = self.workbook.sheetnames
        self.sheet_combo["values"] = names
        if self.active_sheet_name in names:
            self.sheet_combo.set(self.active_sheet_name)
        else:
            self.sheet_combo.set(names[0] if names else "")
            self.active_sheet_name = names[0] if names else None

    def on_sheet_change(self, event=None):
        if not self.workbook: return
        new_sheet = self.sheet_combo.get()
        if new_sheet == self.active_sheet_name: return
        if self.unsaved_changes:
            res = messagebox.askyesnocancel("Unsaved changes", "You have unsaved changes. Save before switching sheets?")
            if res is None:
                self.sheet_combo.set(self.active_sheet_name)
                return
            if res:
                if not self.save_file():
                    self.sheet_combo.set(self.active_sheet_name)
                    return
        self.active_sheet_name = new_sheet
        self._load_active_sheet()

    def _load_active_sheet(self):
        if not self.workbook or not self.active_sheet_name: return
        sheet = self.workbook[self.active_sheet_name]
        self.current_sheet = self.active_sheet_name

        # Header logic
        headers = []
        header_row_idx = None
        for r in sheet.iter_rows(min_row=1, max_row=5):
            values = [cell.value for cell in r]
            if any(v is not None and str(v).strip() != "" for v in values):
                header_row_idx = r[0].row
                headers = [str(cell.value).strip() if cell.value is not None and str(cell.value).strip() != "" else None for cell in r]
                break
        if not headers:
            max_col = sheet.max_column or 1
            headers = [None] * max_col
            header_row_idx = 1
        self.headers = [h if h else f"Column {i+1}" for i,h in enumerate(headers)]

        self.validation_rules = self._infer_validation_rules(self.headers)
        self.current_rules = {rule["name"]: rule for rule in self.validation_rules}
        self.input_order = [rule["name"] for rule in self.validation_rules]
        self._build_input_fields(self.headers)

        self._clear_treeview()
        cols = [f"c{i}" for i in range(len(self.headers))]
        self.tree["columns"] = cols
        for i,h in enumerate(self.headers):
            self.tree.heading(cols[i], text=h, anchor=tk.W)
            self.tree.column(cols[i], width=160, anchor=tk.W)

        start_row = header_row_idx + 1
        rows = []
        # Clear shadows for this sheet to reload fresh
        self.shadow_values = {k:v for k,v in self.shadow_values.items() if k[0] != self.current_sheet}

        for excel_row in range(start_row, sheet.max_row+1):
            row_cells = [sheet.cell(row=excel_row, column=c+1) for c in range(len(self.headers))]
            rowvals = [ (cell.value if cell.value is not None else "") for cell in row_cells ]
            if all(v == "" or v is None for v in rowvals):
                continue
            
            display_row_index = excel_row - start_row + 1
            for col_idx, cell in enumerate(row_cells, start=1):
                raw_value = cell.value
                parsed_value = cell.value
                rounded_flag = False # Initial load assumption
                self._shadow_set(self.current_sheet, display_row_index, col_idx, raw_value, parsed_value, rounded_flag)
            rows.append(rowvals)

        self.all_rows = rows
        for r_idx, row in enumerate(rows, start=1):
            display_row = []
            for col_idx, cell_val in enumerate(row):
                header_name = self.input_order[col_idx]
                rule = self.current_rules.get(header_name, {})
                display_text = format_value_for_display(cell_val, rule=rule)
                display_row.append(display_text)
            row_id = self.tree.insert("", tk.END, values=display_row)
            self._apply_rounded_tags(row_id, display_row)

        self._create_filter_row()
        self._load_user_prefs()

    # -------------------------
    # Rules & Inputs
    # -------------------------
    def _infer_validation_rules(self, headers):
        rules = []
        numeric_keywords = ("qty", "quantity", "number", "count", "age")
        decimal_keywords = ("amount", "price", "rate", "total", "cost", "balance", "value")
        for h in headers:
            h_lower = h.lower() if h else ""
            val_type = "text"
            num_format = None
            is_required_default = True
            duplicate_policy_default = "none"
            if any(k in h_lower for k in decimal_keywords):
                val_type = "numeric"
                num_format = "decimal"
            elif any(k in h_lower for k in numeric_keywords):
                val_type = "numeric"
                num_format = "integer"
            elif "date" in h_lower:
                val_type = "date"
            elif "email" in h_lower:
                val_type = "email"
            if "optional" in h_lower:
                is_required_default = False
            if "id" in h_lower or "code" in h_lower:
                duplicate_policy_default = "strict"
                is_required_default = False
            rule = {
                "name": h,
                "type": val_type,
                "format": num_format,
                "required_var": tk.BooleanVar(value=is_required_default),
                "duplicate_var": tk.StringVar(value=duplicate_policy_default),
                "required": is_required_default,
                "duplicate_policy": duplicate_policy_default
            }
            rules.append(rule)
        return rules

    def _clear_inputs_area(self):
        for w in self.inputs_inner.winfo_children():
            w.destroy()
        self.input_entries.clear()

    def _build_input_fields(self, headers):
        self._clear_inputs_area()
        for idx, header in enumerate(headers):
            rule = self.validation_rules[idx]
            col_frame = ttk.Frame(self.inputs_inner)
            col_frame.grid(row=0, column=idx, padx=6, pady=4)
            lbl = ttk.Label(col_frame, text=header, width=20, anchor="center")
            lbl.pack(side=tk.TOP, fill=tk.X)
            ent = tk.Entry(col_frame, width=20)
            ent.pack(side=tk.TOP, pady=(6,0))
            ent.bind("<Return>", lambda e, i=idx: self._on_enter_pressed(e, i))
            ent.bind("<Tab>", lambda e, i=idx: (self._on_enter_pressed(e, i), "break")[1])
            self.input_entries.append(ent)
            error_var = tk.StringVar(value="")
            error_lbl = ttk.Label(col_frame, textvariable=error_var, foreground="red", anchor="center")
            error_lbl.pack(side=tk.TOP, fill=tk.X)
            ent.error_var = error_var
            
            # Control frame for rules
            control_frame = ttk.Frame(col_frame)
            control_frame.pack(side=tk.TOP, fill=tk.X, pady=(5,0))
            req_chk = ttk.Checkbutton(control_frame, text="Required", variable=rule['required_var'], command=lambda r=rule: self._update_validation_state(r))
            req_chk.pack(anchor=tk.W)
            
            dup_lbl = ttk.Label(control_frame, text="Duplicate Policy:")
            dup_lbl.pack(anchor=tk.W, pady=(2,0))
            ttk.Radiobutton(control_frame, text="None", variable=rule['duplicate_var'], value="none", command=lambda r=rule: self._update_validation_state(r)).pack(anchor=tk.W, padx=10)
            ttk.Radiobutton(control_frame, text="Warn", variable=rule['duplicate_var'], value="warn", command=lambda r=rule: self._update_validation_state(r)).pack(anchor=tk.W, padx=10)
            ttk.Radiobutton(control_frame, text="Strict", variable=rule['duplicate_var'], value="strict", command=lambda r=rule: self._update_validation_state(r)).pack(anchor=tk.W, padx=10)

        self.root.after(100, lambda: self.input_canvas.configure(scrollregion=self.input_canvas.bbox("all")))
        self.reset_to_add_mode()

    def _update_validation_state(self, rule):
        rule["required"] = rule['required_var'].get()
        rule["duplicate_policy"] = rule['duplicate_var'].get()
        self._update_status(f"Updated policy for '{rule['name']}'")

    def _on_enter_pressed(self, event, idx):
        if idx == len(self.input_entries) - 1:
            if self.mode == "add":
                self.add_row_from_inputs()
            elif self.mode == "edit":
                self.update_row_from_inputs()
        else:
            self.input_entries[idx+1].focus_set()

    def _get_existing_column_data(self, col_index):
        if not self.workbook or not self.active_sheet_name:
            return set()
        sheet = self.workbook[self.active_sheet_name]
        existing_values = set()
        start_row = 2
        for row_idx in range(start_row, sheet.max_row+1):
            cell_value = sheet.cell(row=row_idx, column=col_index).value
            if cell_value is not None:
                normalized_value = str(cell_value).strip()
                if normalized_value:
                    existing_values.add(normalized_value.lower())
        return existing_values

    # -------------------------
    # VALIDATION
    # -------------------------
    def validate_inputs(self):
        normalized = []
        strict_messages = []
        warning_messages = []
        is_valid = True
        is_edit_mode = self.mode == "edit"

        for i, entry in enumerate(self.input_entries):
            val = entry.get()
            rule = self.validation_rules[i]
            col_name = rule["name"]
            val_type = rule["type"]
            val_format = rule.get("format")
            required = rule["required"]
            duplicate_policy = rule["duplicate_policy"]

            # Reset Visuals
            try:
                entry.config(bg="white")
            except: pass
            if hasattr(entry, 'error_var'):
                entry.error_var.set("")
            
            val_stripped = str(val).strip()

            # 1. Required Check
            if required and not val_stripped:
                is_valid = False
                strict_messages.append(f"❌ {col_name}: Required.")
                normalized.append(None)
                entry.config(bg="#fbb") # Light Red
                if hasattr(entry, 'error_var'): entry.error_var.set("Required")
                continue
            
            if not val_stripped and not required:
                normalized.append(None)
                continue

            # 2. Duplicate Check
            if duplicate_policy in ("strict","warn"):
                col_index = i + 1
                existing_data = self._get_existing_column_data(col_index=col_index)
                is_original_value = False
                if is_edit_mode:
                    original_val = self.original_editing_values.get(col_index, "__NO_MATCH__")
                    if val_stripped.lower() == original_val:
                        is_original_value = True
                
                if not is_original_value and val_stripped.lower() in existing_data:
                    if duplicate_policy == "strict":
                        is_valid = False
                        strict_messages.append(f"❌ {col_name}: Duplicate value (Strict).")
                        entry.config(bg="#fbb")
                        if hasattr(entry, 'error_var'): entry.error_var.set("Duplicate")
                        normalized.append(val_stripped)
                        continue
                    elif duplicate_policy == "warn":
                        warning_messages.append(f"⚠️ {col_name}: Possible duplicate.")
                        entry.config(bg="#ffdd99") # Orange
                        if hasattr(entry, 'error_var'): entry.error_var.set("Duplicate?")

            # 3. Type parsing & Precision Check
            try:
                if val_type == "text":
                    cleaned = re.sub(r"\s+", " ", val_stripped)
                    normalized.append(cleaned if cleaned else None)

                elif val_type == "numeric":
                    if not is_numeric(val):
                        raise ValueError("Invalid number")
                    
                    # EXCESS PRECISION CHECK
                    if val_format == "decimal":
                        if has_excess_precision(val, decimal_limit=2):
                            warning_messages.append(f"⚠️ {col_name}: Excess precision (>2 decimals). Will be rounded.")
                            entry.config(bg="#e0ccff") # Purple
                            if hasattr(entry, 'error_var'): entry.error_var.set("Rounding alert")

                    normalized.append(normalize_numeric(val, fmt=val_format))

                elif val_type == "date":
                    date_obj = try_parse_date(val)
                    if date_obj is None:
                        raise ValueError("Invalid date")
                    normalized.append(date_obj)

                elif val_type == "email":
                    if not EMAIL_RE.match(val_stripped):
                        raise ValueError("Invalid email")
                    normalized.append(val_stripped)
                else:
                    normalized.append(val_stripped)

            except ValueError as e:
                is_valid = False
                strict_messages.append(f"❌ {col_name}: {e}")
                entry.config(bg="#fbb")
                if hasattr(entry, 'error_var'): entry.error_var.set(str(e))
                normalized.append(val_stripped)

        return is_valid, strict_messages, warning_messages, normalized

    def clear_input_entries(self):
        for ent in self.input_entries:
            try:
                ent.delete(0, tk.END)
                ent.config(bg="white")
            except: pass
            if hasattr(ent, 'error_var'):
                ent.error_var.set("")

    # -------------------------
    # Apply Rounded Tags (Visuals in Tree)
    # -------------------------
    def _apply_rounded_tags(self, row_id, display_row):
        try:
            row_index_display = int(self.tree.index(row_id)) + 1
        except Exception:
            return
        
        for col_idx in range(len(display_row)):
            col_excel = col_idx + 1
            shadow = self._shadow_get(self.current_sheet, row_index_display, col_excel)
            text = display_row[col_idx]
            
            # Apply icon if precision mismatch
            if shadow and shadow.get("rounded_flag"):
                if self.rounded_icon:
                    self.tree.image_create(row_id, column=col_excel - 1, image=self.rounded_icon, sticky="nw")
                
                # Update text prefix
                if not str(text).startswith("▲ "):
                    new_text = f"▲ {text}"
                else:
                    new_text = text
            else:
                # Remove prefix if no longer rounded
                if str(text).startswith("▲ "):
                    new_text = str(text)[2:]
                else:
                    new_text = text
            
            try:
                self.tree.set(row_id, col_idx, new_text)
            except Exception:
                vals = list(self.tree.item(row_id, "values"))
                vals[col_idx] = new_text
                self.tree.item(row_id, values=vals)

    # -------------------------
    # Add / Edit / Delete
    # -------------------------
    def add_row_from_inputs(self):
        if not self.workbook or not self.active_sheet_name:
            messagebox.showwarning("No file", "Open an .xlsx file first.")
            return
        
        is_valid, strict_messages, warning_messages, normalized = self.validate_inputs()
        
        if not is_valid:
            self._update_status(f"Validation failed on {len(strict_messages)} field(s).", "error")
            return

        if warning_messages:
            warning_text = "\n".join(warning_messages)
            prompt = (
                "Notices regarding your input:\n\n"
                f"{warning_text}\n\n"
                "Proceed with adding this record?"
            )
            res = messagebox.askyesno("Input Warning", prompt, icon='warning')
            if not res:
                self._update_status("Cancelled.", "warning")
                return

        # Commit to Workbook & Tree
        sheet = self.workbook[self.active_sheet_name]
        append_row_excel = sheet.max_row + 1
        display_row = []
        display_row_index = len(self.tree.get_children()) + 1

        for idx, entry in enumerate(self.input_entries):
            raw = entry.get()
            parsed = normalized[idx]
            field_name = self.input_order[idx]
            rule = self.current_rules.get(field_name, {})
            display_text = format_value_for_display(parsed, rule=rule)
            
            # Detect rounding for Shadow
            rounded_flag = False
            if parsed is not None and rule.get("format") == "decimal":
                rounded_flag = detect_precision_mismatch(raw, parsed, decimal_places=2)

            col_index_excel = idx + 1
            self._shadow_set(self.current_sheet, display_row_index, col_index_excel, raw, parsed, rounded_flag)
            display_row.append(display_text)
            try:
                sheet.cell(row=append_row_excel, column=col_index_excel).value = parsed
            except Exception:
                sheet.cell(row=append_row_excel, column=col_index_excel).value = parsed

        row_id = self.tree.insert("", tk.END, values=display_row)
        self._apply_rounded_tags(row_id, display_row)
        self.tree.selection_remove(self.tree.selection())
        self.tree.selection_set(row_id)
        self.tree.see(row_id)
        self.unsaved_changes = True
        self._update_status(f"Added new row to '{self.active_sheet_name}'.", "success")
        
        if hasattr(self, 'auto_save_var') and self.auto_save_var.get():
            self.save_file()
            
        self.clear_input_entries()
        if self.input_entries:
            self.input_entries[0].focus_set()

    def update_row_from_inputs(self):
        if not hasattr(self, "editing_item") or not self.editing_item:
            messagebox.showinfo("No selection", "No row selected.")
            return

        is_valid, strict_messages, warning_messages, normalized = self.validate_inputs()
        
        if not is_valid:
            self._update_status("Validation failed.", "error")
            return
            
        if warning_messages:
            warning_text = "\n".join(warning_messages)
            prompt = (
                "Notices regarding your input:\n\n"
                f"{warning_text}\n\n"
                "Update this record anyway?"
            )
            res = messagebox.askyesno("Update Warning", prompt, icon='warning')
            if not res: return

        display_row = []
        display_row_index = int(self.tree.index(self.editing_item)) + 1
        sheet = self.workbook[self.active_sheet_name]
        excel_row = display_row_index + 1

        for idx, entry in enumerate(self.input_entries):
            raw = entry.get()
            parsed = normalized[idx]
            field_name = self.input_order[idx]
            rule = self.current_rules.get(field_name, {})
            display_text = format_value_for_display(parsed, rule=rule)

            rounded_flag = False
            if parsed is not None and rule.get("format") == "decimal":
                rounded_flag = detect_precision_mismatch(raw, parsed, decimal_places=2)

            col_excel = idx + 1
            self._shadow_set(self.current_sheet, display_row_index, col_excel, raw, parsed, rounded_flag)
            display_row.append(display_text)
            try:
                sheet.cell(row=excel_row, column=col_excel).value = parsed
            except Exception:
                pass

        self.tree.item(self.editing_item, values=display_row)
        self._apply_rounded_tags(self.editing_item, display_row)
        self.unsaved_changes = True
        self._update_status(f"Updated row {excel_row-1}.", "success")
        self.reset_to_add_mode()
        
        if hasattr(self, 'auto_save_var') and self.auto_save_var.get():
            self.save_file()

    def on_tree_double_click(self, event):
        selected_item = self.tree.focus()
        if not selected_item: return
        values = self.tree.item(selected_item, "values")
        if not values: return
        self.clear_input_entries()
        display_row_index = int(self.tree.index(selected_item)) + 1
        for idx, entry in enumerate(self.input_entries):
            col_idx = idx + 1
            shadow = self._shadow_get(self.current_sheet, display_row_index, col_idx)
            entry.delete(0, tk.END)
            if shadow and shadow.get("raw") is not None:
                entry.insert(0, shadow.get("raw"))
            else:
                disp = values[idx] if idx < len(values) else ""
                if isinstance(disp, str) and disp.startswith("▲ "):
                    disp = disp[2:]
                entry.insert(0, disp)
        self.editing_item = selected_item
        self.mode = "edit"
        self.original_editing_values = {}
        for i, entry in enumerate(self.input_entries):
            v = entry.get()
            self.original_editing_values[i+1] = str(v).strip().lower()
        self.add_button.config(text="Update Row", command=self.update_row_from_inputs, style="warning.TButton")
        for entry in self.input_entries:
            entry.unbind("<Return>")
            entry.bind("<Return>", lambda e: self.update_row_from_inputs())
        if self.input_entries:
            self.input_entries[0].focus_set()
        self._update_status("Editing existing row...", "warning")

    def reset_to_add_mode(self):
        self.clear_input_entries()
        self.add_button.config(text="➕ Add Row", command=self.add_row_from_inputs, style="success.TButton")
        for idx, entry in enumerate(self.input_entries):
            try:
                entry.unbind("<Return>")
                entry.unbind("<Tab>")
                entry.bind("<Return>", lambda e, i=idx: self._on_enter_pressed(e, i))
                entry.bind("<Tab>", lambda e, i=idx: (self._on_enter_pressed(e, i), "break")[1])
            except Exception:
                pass
        self.mode = "add"
        self.editing_item = None
        self.original_editing_values = {}
        if self.input_entries:
            self.input_entries[0].focus_set()

    def delete_selected_row(self):
        if not self.workbook or not self.active_sheet_name:
            return
        selected_item = self.tree.focus()
        if not selected_item: return
        
        confirm = messagebox.askyesno("Confirm Deletion", "Delete this row?")
        if not confirm: return
        
        try:
            self._flash_tree_row(selected_item, color="#ffcccc", duration=300)
            self.root.after(300, lambda: self._delete_row_after_flash(selected_item))
        except:
            self._delete_row_after_flash(selected_item)

    def _flash_tree_row(self, item_id, color="#ccffcc", duration=800):
        try:
            tag_name = f"flash_{item_id}"
            self.tree.tag_configure(tag_name, background=color)
            self.tree.item(item_id, tags=(tag_name,))
            self.root.after(duration, lambda: self.tree.item(item_id, tags=()))
        except: pass

    def _delete_row_after_flash(self, selected_item):
        try:
            sheet = self.workbook[self.active_sheet_name]
            excel_row_index = self.tree.index(selected_item) + 2
            sheet.delete_rows(excel_row_index, 1)
            
            # Clean up shadow
            self.tree.delete(selected_item)
            self.unsaved_changes = True
            self._update_status("Row deleted.", "success")
            
            if hasattr(self, 'auto_save_var') and self.auto_save_var.get():
                self.save_file()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to delete:\n{e}")

    # -------------------------
    # Save & Prefs
    # -------------------------
    def _flush_shadow_to_workbook(self):
        # We write directly during Add/Edit, but this is a safety flush if needed
        pass 

    def save_file(self):
        if not self.workbook: return False
        if not self.filepath: return self.save_file_as()
        try:
            self.workbook.save(self.filepath)
            self.unsaved_changes = False
            self._update_status(f"Saved: {os.path.basename(self.filepath)}", "success")
            self._save_user_prefs()
            return True
        except Exception as e:
            messagebox.showerror("Save error", str(e))
            return False

    def save_file_as(self):
        filetypes = [("Excel files", "*.xlsx")]
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=filetypes)
        if not path: return False
        self.filepath = path
        try:
            self.workbook.save(self.filepath)
            self.unsaved_changes = False
            self._update_status(f"Saved as: {os.path.basename(self.filepath)}", "success")
            self._save_user_prefs()
            return True
        except Exception as e:
            messagebox.showerror("Save error", str(e))
            return False

    def on_theme_change(self, event=None):
        try:
            Style().theme_use(self.theme_combo.get())
            self._update_status(f"Theme: {self.theme_combo.get()}", "success")
        except: pass

    def _show_temp_warning(self, message, duration=5000):
        # Simplified popup logic
        messagebox.showwarning("Warning", message)

    def _add_hover_effect(self, widget):
        widget.bind("<Enter>", lambda e: widget.configure(cursor="hand2"))
        widget.bind("<Leave>", lambda e: widget.configure(cursor=""))

    def _show_about(self):
        messagebox.showinfo("About", APP_TITLE)

    def _show_help(self):
        help_path = resource_path("help.txt")
        if os.path.exists(help_path):
            os.startfile(help_path) if os.name == 'nt' else None
        else:
            messagebox.showinfo("Help", "No help.txt found.")

    def on_close(self):
        if self.unsaved_changes:
            if messagebox.askyesno("Unsaved changes", "Save before exit?"):
                self.save_file()
        self._save_user_prefs()
        self.root.destroy()

    def _get_prefs_path(self):
        return self.filepath + ".prefs.json" if self.filepath else None

    def _load_user_prefs(self):
        prefs_path = self._get_prefs_path()
        if not prefs_path or not os.path.exists(prefs_path):
            return
        try:
            with open(prefs_path, "r", encoding="utf-8") as f:
                prefs = json.load(f)
            if "theme" in prefs:
                try:
                    self.theme_combo.set(prefs["theme"])
                    Style().theme_use(prefs["theme"])
                except Exception: pass
            if "auto_save" in prefs:
                try: self.auto_save_var.set(bool(prefs["auto_save"]))
                except: pass
            
            sheet_name = getattr(self, "active_sheet_name", "ActiveSheet")
            sheet_prefs = prefs.get("sheets", {}).get(sheet_name, {})
            cols = sheet_prefs.get("columns", {})
            if cols:
                for rule in self.validation_rules:
                    name = rule["name"]
                    if name in cols:
                        try:
                            rule["required_var"].set(bool(cols[name].get("required")))
                            rule["duplicate_var"].set(cols[name].get("duplicate"))
                            self._update_validation_state(rule)
                        except: continue
        except Exception: pass

    def _save_user_prefs(self):
        prefs_path = self._get_prefs_path()
        if not prefs_path: return
        try:
            prefs = {}
            if os.path.exists(prefs_path):
                try:
                    with open(prefs_path, "r", encoding="utf-8") as f:
                        prefs = json.load(f) or {}
                except: prefs = {}
            
            sheet_name = getattr(self, "active_sheet_name", "ActiveSheet")
            prefs.setdefault("sheets", {})
            prefs["theme"] = self.theme_combo.get()
            prefs["auto_save"] = bool(self.auto_save_var.get())
            prefs["sheets"][sheet_name] = {
                "columns": {
                    rule["name"]: {
                        "required": bool(rule["required_var"].get()),
                        "duplicate": rule["duplicate_var"].get()
                    } for rule in self.validation_rules
                }
            }
            with open(prefs_path, "w", encoding="utf-8") as f:
                json.dump(prefs, f, indent=2)
        except: pass

# -------------------------
# Entry point
# -------------------------
def main():
    app_root = Window(title=APP_TITLE, themename="cosmo")
    app = DynamicExcelApp(app_root)
    app_root.mainloop()

if __name__ == "__main__":
    main()