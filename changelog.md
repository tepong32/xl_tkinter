# Changelog

## [4.1.2] - 2025-10-17
### 🚀 Added
- test of the new version manager

# 🧾 Project Changelog — Excel-Style Tkinter App

All notable changes to this project will be documented here.
This file follows the [Keep a Changelog](https://keepachangelog.com/en/1.1.0/) format.

---

## [v4.2.0] - 2025-10-17

### 🚀 Added
- **Column-Based Duplicate Checking (Uniqueness)**
  - Introduced **mandatory uniqueness checks** for designated columns (e.g., fields containing `(Unique)` or `ID` in the header).
  - Prevents duplicate entries from being added or created during edits.
- **Type-Specific Validation**
  - Implemented strict type validation for **Numeric**, **Date**, and **Email** fields based on header keywords.
  - Fields will now reject invalid input formats (e.g., non-numeric text, improperly formatted dates).
- **Edit Mode Exclusion for Uniqueness**
  - The uniqueness check intelligently **excludes the original cell's value** when a row is being edited, preventing false positive errors.

### 💡 Improved
- **Validation Rules Inference**
  - The rule inference (`_infer_validation_rules`) now correctly identifies and marks fields as **Required (R)** and/or **Unique (U)** directly in the input field labels.
- **User Feedback**
  - Validation messages are now shown in a specific error label beneath each input field, in addition to the warning message box, for better field-level feedback.
- **Enter Key Behavior**
  - The `<Return>` key binding logic is now unified to either **Add Row** (in 'add' mode) or **Update Row** (in 'edit' mode) from any input field.

### 🧠 Technical Notes
- Introduced `_get_existing_column_data` helper for efficient set-based checking of unique values across the spreadsheet.
- The `original_editing_values` dictionary is used to safely manage uniqueness checks in 'edit' mode.
- The `validate_inputs` method now incorporates the full suite of type and uniqueness checks.

---

## [v4.1.1] - 2025-10-17

### 💡 Improved
- Double-Click Edits and Delete Rows now working.
- Updated visuals: buttons, status indicators, etc.

---

## [v4.1.0] - 2025-10-16

### 🚀 Added
- Added **edit rows** and **auto-save** features.

### 🧹 Changed
- Text validation still applies on edits.
- Highlights on new row additions need to be re-implemented.
- UI re-work for indicators and buttons needed.

---

## [v4.0.2] - 2025-10-16
- Trying patch versioning.

## [v4.0.1] - 2025-10-16
- Testing out python version manager.

## [v4.0.0] - 2025-10-16
- Testing out python version manager.

---

## [v2.0.0] - 2025-10-14

### 🚀 Added
- **Dynamic Excel-Driven Data Entry**
  - Automatically generates input fields based on the active sheet’s headers — no more fixed layouts.
- **Sheet Switching Support**
  - Dropdown (combobox) added to switch between sheets in the loaded workbook.
  - The entry form and data table rebuild dynamically per sheet.
- **Startup File Selection**
  - Prompts the user to choose an Excel file (`.xlsx`) at launch, allowing flexible templates.
- **Treeview Data Display**
  - Displays current sheet data below the input area for quick review.
  - Newly added rows are auto-highlighted for visual confirmation.
- **Save-on-Exit Prompt**
  - Users are now asked to save before closing — mimicking Excel’s behavior.
- **Input Cleanup & Validation Feedback**
  - Inputs automatically trim leading, trailing, and multiple in-between spaces.
  - Invalid fields highlight red (`#ffe6e6`), optional blanks (e.g., “Notes/Remarks”) yellow (`#fff7cc`), and valid fields reset to white.
  - Real-time validation runs before committing data to Excel.

### 💡 Improved
- **Adaptive Layout**
  - The UI now flows top-to-bottom: entry fields first, data table below, creating a natural “press Enter to commit” workflow.
- **Automatic Sheet Reconfiguration**
  - Switching sheets refreshes field definitions and validation rules on the fly.
- **Code Structure Refinement**
  - Introduced `clean_spaces()` helper and reorganized validation logic for maintainability.
- **Theme System**
  - Light/Dark theme toggle preserved and fully compatible with new layout.

### 🧠 Technical Notes
- Core Excel handling powered by **OpenPyXL**.
- Input validation, highlight logic, and theme behavior retained from earlier versions.
- Final column (usually “Notes” or “Remarks”) is optional by design.
- Lays groundwork for autosave, editable rows, and import/export features in future versions.

### 🔜 Planned
- Export to CSV / JSON / SQL
- Auto-save intervals and recovery
- “New Excel File” creation wizard

### 🧩 Version Summary

| Component | Status |
| :--- | :--- |
| Core Excel I/O | ✅ Stable |
| Dynamic Input UI | ✅ Implemented |
| Theme Toggle | ✅ Preserved |
| Save-on-Exit | ✅ Implemented |
| Sheet Switcher | ✅ Implemented |
| Validation | **✅ Enhanced** |
| Export Options | ⏳ Planned |
| Auto-Save | ✅ Implemented |
| Type-Specific Inputs | **✅ Implemented** |

---

## [v0.8.2] - 2025-10-13

### ✨ Added
- Theme Selection Combobox for instant Light/Dark mode switching.
- Top bar with theme indicator for better UX.

### 💡 Improved
- Header layout and toggler placement.
- Grid-based entry layout maintained.

### 🧹 Changed
- Temporarily disabled validation highlighting (for refactor).
- Streamlined imports and comments.

---

## [v0.8.1] - 2025-10-11

### 🧩 Added
- Base Excel integration using `openpyxl`.
- Auto-create file with headers if missing.
- Load and insert data directly from/to sheet.
- “Add Record” button with simple input form.

### 🎨 Improved
- Centered main window (1280×960) with min-size limits.
- Applied `ttkbootstrap` themes.
- Reversed display order (latest entries first).

---

## [v0.8.0] - 2025-10-09

### 🚀 Initial Commit
- Base Tkinter + ttkbootstrap structure.
- Implemented form layout, data table (Treeview), and Excel I/O foundation.
