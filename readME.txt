# 🧮 Dynamic Excel Data Entry App (Tkinter + ttkbootstrap)

A modern **Excel-style data entry companion** built with **Tkinter** and **ttkbootstrap**, designed for people who want the familiarity of spreadsheets — with the validation, structure, and polish of a real app.  

> ✨ Perfect for office staff, students, and small-scale data collectors who work with Excel templates daily.

---

## 🚀 Key Features

### 🗂 Excel Integration
- Opens and edits `.xlsx`, `.xlsm`, `.xlsb`, `.xls`, and `.ods` files  
- Automatically converts non-`.xlsx` formats for safe editing  
- Preserves headers and data types intelligently  

### 💡 Smart Validation
- Detects data types based on header names (`Date`, `Email`, `Amount`, etc.)  
- Auto-formats numbers:
  - `amount`, `price`, `rate`, `total`, `cost`, `balance` → rounded to **2 decimals**
  - `qty`, `age`, `count`, `number` → treated as **integers**
- Real-time input feedback (color + message per field)
- Inline duplicate detection with **None / Warn / Strict** policies  
- “Required” toggles per column for easy setup  

### ⚙️ Excel-Like Workflow
- Keyboard shortcuts for rapid entry:
  - **Ctrl+O** → Open file  
  - **Ctrl+S** / **Ctrl+Shift+S** → Save / Save As  
  - **Ctrl+N** → Clear input fields  
  - **Ctrl+D** → Duplicate selected row  
  - **Ctrl+Shift+D** → Delete selected row  
  - **Ctrl+Shift+I** → Insert blank row below selection  
  - **F2** → Edit selected row  
  - **Esc** → Cancel edit / Reset to add mode  
  - **Ctrl+Q** → Quit the app  
- Auto-increment for ID-like fields (`ID001 → ID002`, `Ref10 → Ref11`)
- Visual row feedback (green for insert/duplicate, red for delete)
- Auto-save option on Add/Edit/Delete

### 🎨 Polished Interface
- Built with **ttkbootstrap** for a clean modern look  
- Theme toggle with all available ttkbootstrap themes (e.g. Cosmo, Darkly, Flatly, etc.)
- Responsive layout with scrollable inputs  
- Sheet selector for multi-sheet workbooks  
- Status bar for real-time updates

---

## 🧠 How to Use

1. **Prepare your Excel template**
   - Row 1 should contain clear headers (e.g. `Name`, `Email`, `Amount`, `Date`)
   - Each header determines its validation type automatically.
2. **Launch the app**
   ```bash
   python test.py
   ```
3. **Open your Excel file**
   - Choose *Yes* when prompted, or use **Ctrl+O**.
4. **Enter or edit data**
   - Use **Tab** or **Enter** to move through fields.
   - Press **Enter** in the last field to add the row.
   - Double-click a row to edit, then press **Enter** to save changes.
5. **Save your work**
   - Press **Ctrl+S** anytime, or enable **Auto-Save on Add/Edit/Delete**.
6. **Navigate rows quickly**
   - Duplicate or insert rows as needed using keyboard shortcuts.

---

## ✅ Best Practices

| Task | Recommendation |
|------|----------------|
| Header naming | Use clear, descriptive names (e.g., “Amount”, “Email Address”, “Date of Birth”). |
| Number formatting | Use “Amount”, “Total”, or “Price” for 2-decimal rounding; “Qty” or “Age” for integers. |
| Duplicate control | Mark important ID fields as **Strict**, and optional notes as **None**. |
| Required fields | Enable “Required” for any must-fill columns. |
| Saving | Always save after major edits; use Auto-Save for safety. |
| Backups | Keep a copy of your `.xlsx` before bulk edits. |
| Themes | Switch themes anytime from the dropdown on the toolbar. |

---

## 🧩 Requirements

- Python 3.10+  
- Install dependencies:
  ```bash
  pip install ttkbootstrap openpyxl pandas pyxlsb odfpy xlrd
  ```
- On Linux, make sure Tkinter matches your Python version:
  ```bash
  sudo apt-get install python3.11-tk
  ```

---

## 🧾 Credits

- **Theme** → [rdbende/Forest-ttk-theme](https://github.com/rdbende/Forest-ttk-theme)  (initial app build but it now uses ttkbootstrap)
- **Tutorial Reference** → [Tkinter + ttkbootstrap YouTube Guide](https://www.youtube.com/watch?v=8m4uDS_nyCk)  
- Built with ❤️ using Python, ttkbootstrap, and openpyxl  

---

## 📜 License

This project is open for personal and educational use.  
For commercial use or redistribution, please credit the author.

---

## 🧭 Version History

See [`CHANGELOG.md`](./changelog.md) for full version details.
