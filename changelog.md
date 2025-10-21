# Changelog

## [5.1.0] - 2025-10-21
### 🚀 Added
- Expanded data-entry capabilities: Excel-like shortcuts, row duplication/insertion with animation, and header-aware numeric formatting.

## [4.3.0] - 2025-10-20
### 🚀 Added
- added togglers for required fields and duplicate policies

## [4.2.0] - 2025-10-20
### 🚀 Added
- **Duplicate Checker Overhaul**
  - Introduced per-header duplicate control options.
  - Users can now choose whether duplicates should be ignored or allowed based on selected column headers.
  - Upcoming feature: checkbox-based header selection for duplicate checking (replacing reliance on Excel’s headers).

---

## [4.1.3] - 2025-10-17
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

---

## [4.1.2] - 2025-10-17
### 🚀 Added
- Integrated initial test run for the new version manager system.

---

## [4.1.1] - 2025-10-17
### 💡 Improved
- Fixed double-click edit and row deletion issues.
- Polished UI (buttons, status indicators, and colors).

---

## [4.1.0] - 2025-10-16
### 🚀 Added
- Introduced **edit rows** and **auto-save** functionality.

### 🧹 Changed
- Validation now applies during edits.
- Highlighting for newly added rows pending re-implementation.
- UI cleanup for future indicator updates.

---

## [4.0.2] - 2025-10-16
### 🧪 Patch
- Testing patch version increment behavior.

## [4.0.1] - 2025-10-16
### 🧪 Patch
- Initial test for Python-based version manager integration.

## [4.0.0] - 2025-10-16
### 🚀 Major
- Migration to Python version manager workflow.

---

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

## [0.8.2] - 2025-10-13
### ✨ Added
- Theme selection combobox for light/dark modes.
- Top bar with active theme indicator.

### 💡 Improved
- Header layout and toggler placement.
- Grid-based entry layout cleanup.

### 🧹 Changed
- Temporarily disabled validation highlighting for refactor.

---

## [0.8.1] - 2025-10-11
### 🧩 Added
- Base Excel integration using **OpenPyXL**.
- Auto-create Excel file with default headers.
- Record insertion and display in Treeview.

### 🎨 Improved
- Centered window (1280×960) with min-size limits.
- Applied **ttkbootstrap** themes.
- Reversed row display order (latest first).

---

## [0.8.0] - 2025-10-09
### 🚀 Initial Commit
- Tkinter + ttkbootstrap structure.
- Core Excel I/O and table integration.
- Data entry form + Treeview foundation.

---

# 🧾 Project Changelog — Excel-Style Tkinter App
All notable changes to this project will be documented here.
This file follows the [Keep a Changelog](https://keepachangelog.com/en/1.1.0/) format.
