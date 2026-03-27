# Changelog

All notable changes to ACLens will be documented in this file.
Format based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).

---

## [0.2.1-alpha] — 2026-03-27

### Fixed
- `switch` statement spacing caused `UnexpectedToken` parse error on PowerShell 5.1 in `Get-SPExternalSharing`

### Changed
- Version label added to main window header (top left, next to "ACLens")
- Subtitle updated to "NTFS & SharePoint Permission Analyzer"

---

## [0.2.0-alpha] — 2026-03-27

### SharePoint Online support

#### Added
- **SharePoint tab** in the main GUI alongside the NTFS tab
- **Setup Wizard** with two modes:
  - **Automatic** (Device Code Flow) — ACLens creates the Azure AD App Registration automatically; requires one-time sign-in with an Application Administrator account
  - **Manual** — enter Tenant ID, Client ID and Client Secret from an existing App Registration
- Encrypted credential storage in `sp_credentials.xml` using Windows DPAPI
- **SharePoint scanner** via Microsoft Graph API:
  - Site-level permissions
  - Document Libraries and Lists
  - Recursive folder tree with per-folder unique permission detection
  - SharePoint Groups & Members
  - External Sharing status
- **SP HTML report** — source-badged, collapsible library/folder tree, unique vs. inherited permission detection, sharing link detection
- **SP JSON snapshot** auto-saved alongside every SP HTML report
- **Compare Scans** window extended with SharePoint Compare tab:
  - Select SP baseline JSON + current site URL
  - Runs a live scan and generates SP diff report
  - Shows added/removed folders and exact permission changes per folder
- [Manual SharePoint Setup Guide](docs/manual-sp-setup.md) added

#### Changed
- Main window now has NTFS / SharePoint tab bar
- Start Analysis button routes to correct scanner based on active tab
- Compare Scans window now has NTFS Compare / SharePoint Compare tabs
- Subtitle updated to reflect SP support

---

## [0.1.0-alpha] — 2026-03-14

### Initial public alpha release

#### Added
- Recursive NTFS permission scan with configurable maximum depth
- Change detection — highlights folders where permissions differ from parent
- Interactive HTML report with filter, search, expand/collapse and print
- Collapsible legend in HTML report explaining all status codes and permission types
- Automatic JSON snapshot export alongside every HTML report
- **Compare Scans** — NTFS diff showing added/removed folders and exact per-folder permission changes (added/removed rules, owner, inheritance)
- Full Windows Forms GUI with progress bar, live status and cancel support
- **Open Last Report** button — reopens last generated HTML report
- **Changelog** button — shows version history in a modal dialog
- **Compare Scans** button in footer
- Admin privilege check on startup with custom-styled dialog and restart-as-admin option
- Output files (HTML + JSON) saved to script folder (`$PSScriptRoot`) by default
- Color scheme based on DiskLens palette (`#1E1E2E` base, `#7C3AED` accent)

#### Known Issues at release
- GUI not DPI-aware on very high-resolution displays
- No CSV / Excel export
- No scheduled scan support
