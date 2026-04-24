# Changelog

All notable changes to ACLens will be documented in this file.
Format based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).

---

## [0.4.0-beta] — 2026-04-16

### GUI & UX

#### Added
- **DE/EN language toggle** — button in header, switches entire GUI and HTML reports between German and English
- **Light mode** for main GUI — toggle button top right of header
- Changelog window is now resizable and theme-aware (respects light/dark mode)
- SharePoint Setup Wizard now respects light/dark mode
- Placeholder text in manual setup fields clears on focus and restores on blur
- Language-aware documentation links (DE → German guide, EN → English guide)

#### Fixed
- All footer buttons now correctly show white text (FlatStyle enum fix, UseVisualStyleBackColor)
- Cancel button always visible — no longer disabled on startup
- Duplicate `New-CompareHTMLReport` function removed
- All `FlatStyle = "Flat"` string assignments replaced with proper enum throughout
- Self-referencing entries in EN string table fixed (wizard was blank in EN)
- Inline `if` expressions on property assignments replaced with variables (PS 5.1 fix)
- Secondary buttons unified to consistent gray in both light and dark mode
- SP Setup Wizard panel widths corrected to match window size
- Footer buttons centered horizontally and vertically

### HTML Report

#### Added
- **Dark/Light mode toggle** — button top right, preference saved in localStorage
- **Paginated view** — 25/50/100/All folders per page with prev/next navigation
- Only Root and Changed folders expanded by default
- German translation of all report strings when DE mode active

---

## [0.3.1-alpha] — 2026-04-16

### Fixed
- `switch` statement spacing caused `UnexpectedToken` on PowerShell 5.1
- Version label added to main window header
- Report title corrected to "ACLens — NTFS Permission Report"

---

## [0.3.0-alpha] — 2026-04-16

### HTML Report Redesign
- Paginated view instead of infinite scroll
- Dark/Light mode toggle with localStorage
- Smarter default view (Root + Changed expanded, rest collapsed)
- Expand/Collapse affects current page only
- Improved permission table layout

---

## [0.2.1-alpha] — 2026-03-27

### Fixed
- `switch` statement spacing fix for PS 5.1
- Version badge added to header
- Subtitle updated

---

## [0.2.0-alpha] — 2026-03-27

### SharePoint Online Support
- SharePoint tab in main GUI
- Setup Wizard: Device Code Flow (automatic App Registration)
- Setup Wizard: Manual mode with Tenant ID / Client ID / Secret
- Encrypted credential storage (DPAPI)
- SP Scanner: Sites, Libraries, Folders, Groups, External Sharing
- SP HTML report with collapsible library/folder tree
- SP JSON snapshot auto-saved
- Compare Scans extended: NTFS + SharePoint tabs
- SP Diff report: added/removed folders, changed permissions
- Manual SP setup guides added (EN + DE)

---

## [0.1.0-alpha] — 2026-03-14

### Initial Release
- Recursive NTFS permission scan with configurable depth
- Change detection vs. parent folder
- Interactive HTML report (filter, search, expand/collapse, print)
- Collapsible legend
- JSON snapshot export
- Compare Scans (NTFS diff)
- Windows Forms GUI with progress bar and cancel
- Admin check on startup
- Output saved to script folder
