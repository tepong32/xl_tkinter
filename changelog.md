# 🧾 Project Changelog — Excel-Style Tkinter App


All notable changes to this project will be documented here.  
This file follows the [Keep a Changelog](https://github.com/tepong32/xl_tkinter/en/1.1.0/) format.

## [v2.0.0] - 2025-10-14  
### ✨ Added  
- **Dynamic Excel-Driven Data Entry:**  
  The app now automatically generates input fields based on the headers of the selected Excel sheet. No more fixed layouts — each sheet defines its own structure.

- **Sheet Switching Support:**  
  Added a dropdown (combobox) to switch between sheets in the loaded workbook.  
  The input area and data preview dynamically rebuild to match the selected sheet.

- **Startup File Selection:**  
  The app now starts cleanly and prompts the user to select an Excel file (.xlsx) to work on. This ensures flexibility and allows any existing workbook to be used as a data template.

- **Treeview Data Display:**  
  Introduced a bottom-pane Treeview that displays the sheet’s current data for easy reference.  
  The upper pane is reserved for dynamic entry fields and the “Add Row” button.

- **Save-on-Exit Prompt:**  
  Users are now prompted to save changes when closing the app — just like Excel’s native behavior.

### 🌗 Improved  
- **Theme Toggle Preserved:**  
  The light/dark theme toggle remains fully functional, retaining your preferred look.

- **Layout Organization:**  
  - Top section: Dynamic data entry form  
  - Bottom section: Data table view  
  - Improved readability and resizing behavior  

- **Refactored Structure:**  
  Streamlined helper functions for loading files, switching sheets, and generating UI elements.

- **Validation Integration:**  
  Existing input validation and highlight logic were adapted to work dynamically with generated fields.

### 🧠 Technical Notes  
- Based on **OpenPyXL** for Excel handling.  
- Preserves existing validation, highlighting, and theme logic.  
- Modular design allows future extensions like export options, auto-save, and per-column input types.

### 🔜 Planned  
- Export to CSV / JSON / SQL  
- Auto-save intervals  
- Type-specific input validations (numeric, date, etc.)  
- “New Excel File” creation wizard

**Commit Message:**  



### 🧩 Version Summary  

| Component            | Status        |
| -------------------- | ------------- |
| Core Excel I/O       | ✅ Stable      |
| Dynamic Input UI     | ✅ Implemented |
| Theme Toggle         | ✅ Preserved   |
| Save-on-Exit         | ✅ Implemented |
| Sheet Switcher       | ✅ Implemented |
| Validation           | ✅ Integrated  |
| Export Options       | ⏳ Planned     |
| Auto-Save            | ⏳ Planned     |
| Type-Specific Inputs | ⏳ Planned     |


---

## [v0.8.2] - 2025-10-13
### ✨ Added
- Theme Selection Combobox for instant Light/Dark mode switching.
- Top bar with current theme indicator for better UX.

### 💡 Improved
- Clean header layout with more intuitive placement of theme toggler.
- Preserved grid-based entry layout for familiarity.

### 🧹 Changed
- Temporarily disabled validation highlighting (to be reworked in refactor).
- Streamlined imports and widget setup; added inline comments.

---

## [v0.8.1] - 2025-10-11
### 🧩 Added
- Base Excel integration using `openpyxl`.
- Auto-create Excel file with headers if missing.
- Load and insert data directly from/to Excel sheet.
- “Add Record” button with simple input form.

### 🎨 Improved
- Centered main window (1280×960) and added min-size limits.
- Used `ttkbootstrap` themes for modernized UI.
- Reversed data display (latest entries appear first).

---

## [v0.8.0] - 2025-10-09
### 🚀 Initial Commit
- Created base Tkinter + ttkbootstrap project structure.
- Implemented form layout, data table (Treeview), and Excel I/O foundation.

---

### 🗓 Upcoming
- [ ] Modularize logic into UI, Validation, and ExcelHandler classes.
- [ ] Reimplement smart validation with visual cues.
- [ ] Add Export/Import sheets, auto-save toggle, and status footer.
- [ ] Improve responsiveness for smaller window sizes.

---