# Changelog

## [3.3.0] - 2025-10-22
### 🚀 Added
- **QoL Validation Controls**
  - Each header now has “Required” checkboxes and **None/Warn/Strict** duplicate policies.
  - Dynamic tooltips, hover effects, and inline error labels for clarity.
- **Adaptive Help System**
  - Scrollable help card that loads `help.txt` directly in-app.
  - Added ❓ Help button in toolbar with tooltip and hover animation.
- **Automatic Validation Inference**
  - Detects numeric, decimal, date, and email columns automatically from headers.
  - Rounds decimals to 2 places and keeps integers clean.

### 💡 Improved
- Refined **status bar** with color-coded feedback and fade-out animation.
- Unified hover and tooltip styling for all toolbar controls.
- Centralized Auto-Save + Theme + Help into one right-aligned toolbar cluster.
- Polished Treeview editing and focus logic for seamless Excel-like workflow.

### ⚙️ Keyboard & UX Enhancements
- Excel-style shortcuts:
  - `Ctrl+O` → Open  
  - `Ctrl+S` / `Ctrl+Shift+S` → Save / Save As  
  - `Ctrl+N` → Clear inputs  
  - `Ctrl+D` → Duplicate  
  - `Ctrl+Shift+D` → Delete  
  - `Ctrl+Shift+I` → Insert row  
  - `F2` → Edit row  
  - `Esc` → Reset to Add mode  
  - `Ctrl+Q` → Quit
- Automatic ID-like increment (`ID001 → ID002`, `Ref10 → Ref11`)
- Auto-save after Add/Edit/Delete/Insert/Duplicate when enabled

### 🧠 Technical Notes
- Introduced `_get_existing_column_data` for efficient uniqueness checks.
- Added `original_editing_values` tracking to skip duplicates during edit.
- Enhanced `_fade_status()` for smooth color transitions.
- Fully compatible with `.xlsx`, `.xlsm`, `.xlsb`, `.xls`, `.ods`.

---

## [3.2.0] - 2025-10-22
### 🚀 Added
- Help Card reading from `help.txt` (scrollable, themed window)
- ❓ Help button beside Theme selector with tooltip + hover

### 🧠 Improved
- Grouped Auto-Save, Theme, Help controls in right-aligned toolbar cluster
- Introduced unified tooltip and hover behavior for all toolbar controls
- Replaced old status bar with color-coded, right-aligned version featuring fade-out transitions
- Minor styling cleanup and readiness for Markdown help support

---

## [3.1.0] - 2025-10-21
### 💡 Enhanced
- Unified Treeview focus and selection logic  
- Added Excel-like keyboard shortcuts for insert, delete, duplicate  
- Added color feedback for all row operations  
- Improved auto-focus and clarity after deletion or duplication

---

## [3.0.0] - 2025-10-21
### ✨ Added
- Excel-style workflow with row insert/delete/duplicate  
- Header-aware numeric formatting and validation inference  
- Multi-format Excel loading (.xlsx, .xlsm, .xlsb, .xls, .ods)  
- Animated row feedback and robust autosave

### 🧠 Technical
- Uses OpenPyXL + pandas + ttkbootstrap  
- Lays foundation for future Markdown help & export functions  

---

# 🧾 Project Changelog — Dynamic Excel Data Entry App
All notable changes are documented here following [Keep a Changelog](https://keepachangelog.com/en/1.1.0/) format.
