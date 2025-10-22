# Changelog

## [3.2.0] - 2025-10-22
### 🚀 Added
- Help Card reading from help.txt (scrollable, themed window)
- ❓ Help button beside Theme selector with tooltip + hover

### 🧠 Improved
- Grouped Auto-Save, Theme, Help controls in right-aligned toolbar cluster
- Introduced unified tooltip and hover behavior for all toolbar controls
- Replaced old status bar with color-coded, right-aligned version featuring fade-out transitions
- Minor styling cleanup and readiness for future Markdown help support

## [3.1.0] - 2025-10-21
### 🚀 Added
💡 Enhanced Treeview focus & selection consistency
### 💡 Enhanced
- Unified Treeview selection and focus handling for all row operations
- Eliminated mismatch between visible highlight and active record
- Improved user clarity and precision during row duplication, insertion, and deletion
- Added automatic focus shifting after deletion (next or previous row)
- Updated main and backup versions to current build

### ⚙️ Keyboard Shortcut Improvements
- Ctrl+D → Duplicate selected row (new row now visually highlighted)
- Ctrl+Shift+D → Delete selected row with visual feedback
- Ctrl+Shift+I → Insert blank row below selection with proper focus

## [3.0.0] - 2025-10-21
### ✨ Added
- Excel-style keyboard shortcuts for a smoother workflow  
  (`Ctrl+O`, `Ctrl+S`, `Ctrl+Shift+S`, `Ctrl+N`, `Ctrl+D`, `Ctrl+Shift+D`, `Ctrl+Shift+I`, `F2`, `Esc`, `Ctrl+Q`)
- Insert, delete, and duplicate row functionality with animated visual feedback  
  (green flash for insert, red flash for delete)
- Duplicate row feature that automatically increments ID-like fields (e.g., `ID001 → ID002`)
- Header-aware numeric formatting:
  - `amount`, `price`, `rate`, `total`, `cost`, and `balance` fields now round to **two decimals**
  - Other numeric fields default to **whole numbers** but accept floats
- Universal spreadsheet support:
  - Now opens `.xlsx`, `.xlsm`, `.xlsb`, `.xls`, and `.ods` formats
  - Non-`.xlsx` files are auto-converted for safe editing and saving
- Enhanced status updates, error handling, and UX polish for better data-entry flow

### 🧠 Improved
- Validation logic now respects header context and numeric type inference  
- Auto-save works seamlessly across all row operations (add/edit/delete/duplicate/insert)
- General code cleanup for stability and maintainability

### 🪄 Notes
- This release focuses on making data entry more natural and Excel-like while enforcing validation consistency.  
- Backward compatible — existing `.xlsx` files work without changes.

## [2.7.0] - 2025-10-20
### 🚀 Added
- **Duplicate Checker Overhaul**
  - Introduced per-header duplicate control options.
  - Users can now choose whether duplicates should be ignored or allowed based on selected column headers.
  - Upcoming feature: checkbox-based header selection for duplicate checking (replacing reliance on Excel’s headers).

## [2.6.0] - 2025-10-17
### 🚀 Added
- **Column-Based Duplicate Checking (Uniqueness)**
  - Enforced uniqueness on columns tagged as `(Unique)` or containing “ID”.
  - Prevents duplicate entries on add/edit operations.
- **Type-Specific Validation**
  - Added field validation for numeric, date, and email types inferred from header keywords.
- **Edit Mode Uniqueness Exclusion**
  - Editing now ignores the current cell’s original value when checking duplicates.

### 💡 Improved
- Inferred rules (`_infer_validation_rules`) now mark fields as **Required (R)** or **Unique (U)** directly in the input labels.
- Inline validation messages appear under each input field for better clarity.
- Unified Enter key logic to automatically add or update rows based on the mode.

### 🧠 Technical Notes
- Added `_get_existing_column_data` helper for efficient set-based unique checks.
- Introduced `original_editing_values` tracking for safe edit validation.
- Enhanced `validate_inputs` for full type + uniqueness enforcement.

## [2.5.0] - 2025-10-17
### 🚀 Added
- Integrated initial test run for the new version manager system.

## [2.4.0] - 2025-10-17
### 💡 Improved
- Fixed double-click edit and row deletion issues.
- Polished UI (buttons, status indicators, and colors).

## [2.3.0] - 2025-10-16
### 🚀 Added
- Introduced **edit rows** and **auto-save** functionality.

### 🧹 Changed
- Validation now applies during edits.
- Highlighting for newly added rows pending re-implementation.
- UI cleanup for future indicator updates.

## [2.0.2] - 2025-10-16
### 🧪 Patch
- Testing patch version increment behavior.

## [2.0.1] - 2025-10-16
### 🧪 Patch
- Initial test for Python-based version manager integration.

## [2.0.0] - 2025-10-14
### 🚀 Added
- **Dynamic Excel-Driven Data Entry**
  - Auto-generates input fields from sheet headers.
- **Sheet Switching Support**
  - Dropdown to change active sheet dynamically.
- **Startup File Selection**
  - Prompts user for Excel file at launch.
- **Treeview Data Display**
  - Highlights new entries for better visibility.
- **Save-on-Exit Prompt**
  - Asks to save before exiting.
- **Input Cleanup & Validation Feedback**
  - Real-time validation with colored field feedback.

### 💡 Improved
- Adaptive UI layout and validation structure.
- Refactored helper methods for maintainability.
- Theme toggle and light/dark compatibility.

### 🧠 Technical Notes
- Uses **OpenPyXL** for all Excel I/O.
- Establishes groundwork for autosave, edit, and export features.

---

## [0.1.2] - 2025-10-13
### ✨ Added
- Theme selection combobox for light/dark modes.
- Top bar with active theme indicator.

### 💡 Improved
- Header layout and toggler placement.
- Grid-based entry layout cleanup.

### 🧹 Changed
- Temporarily disabled validation highlighting for refactor.

---

## [0.1.1] - 2025-10-11
### 🧩 Added
- Base Excel integration using **OpenPyXL**.
- Auto-create Excel file with default headers.
- Record insertion and display in Treeview.

### 🎨 Improved
- Centered window (1280×960) with min-size limits.
- Applied **ttkbootstrap** themes.
- Reversed row display order (latest first).

---

## [0.1.0] - 2025-10-09
### 🚀 Initial Commit
- Tkinter + ttkbootstrap structure.
- Core Excel I/O and table integration.
- Data entry form + Treeview foundation.

---

# 🧾 Project Changelog — Excel-Style Tkinter App
All notable changes to this project will be documented here.
This file follows the [Keep a Changelog](https://keepachangelog.com/en/1.1.0/) format.
