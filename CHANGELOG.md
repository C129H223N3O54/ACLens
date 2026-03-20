# Changelog

All notable changes to ACLens will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).

---

## [0.1.0-alpha] — 2026-03-14

### Initial public alpha release

#### Added
- Recursive NTFS permission scan with configurable maximum depth
- Change detection — highlights folders where permissions differ from parent
- Interactive HTML report with filter, search, expand/collapse and print support
- Collapsible legend in HTML report explaining all permission types, scopes and inheritance options
- Automatic JSON snapshot export alongside every HTML report for baseline comparison
- **Compare Scans** feature — diff two scan results showing added/removed folders and exact permission changes per folder (new rules, removed rules, owner changes, inheritance changes)
- Full Windows Forms GUI — no command line knowledge required
- Progress bar with live status and cancel support during scan
- **Open Last Report** button — reopens the most recently generated HTML report
- **Changelog** button — shows version history in a modal dialog
- **Compare Scans** button — opens the comparison window
- Admin privilege check on startup with custom-styled dialog and option to restart as Administrator
- Output files (HTML + JSON) saved to the script folder by default
- Color scheme based on DiskLens design palette (`#1E1E2E` base, `#7C3AED` accent)

#### Known Issues
- GUI is not DPI-aware on very high-resolution displays
- Compare Scans always uses unlimited depth for the live scan regardless of the main scan depth setting
- No CSV or Excel export
- No scheduled/automated scan support

---

*Older entries will be added as development continues.*
