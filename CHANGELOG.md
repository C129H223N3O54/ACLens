# Changelog

All notable changes to ACLens will be documented in this file.
Format based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).

---

## [0.1.0-alpha] — 2026-03-14

### Initial public alpha release

#### Added
- Recursive NTFS permission scan with configurable maximum depth
- Change detection — highlights folders where permissions differ from parent
- Interactive HTML report with filter, search, expand/collapse and print
- Collapsible legend in HTML report
- Automatic JSON snapshot export alongside every HTML report
- **Compare Scans** — diff two snapshots showing added/removed folders and exact permission changes per folder (rules, owner, inheritance)
- Full Windows Forms GUI — no command line required
- Progress bar with live status and cancel support
- **Open Last Report** button
- **Changelog** button
- **Compare Scans** button
- Admin privilege check on startup with custom-styled dialog and restart-as-admin option
- Output files saved to script folder by default (`$PSScriptRoot`)
- Color scheme based on DiskLens palette (`#1E1E2E` base, `#7C3AED` accent)

#### Known Issues
- GUI not DPI-aware on high-resolution displays
- Compare Scans uses unlimited depth regardless of main scan depth setting
- No CSV / Excel export
- No scheduled scan support
