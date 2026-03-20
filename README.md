# ACLens

> **NTFS Permission Analyzer & Reporter for Windows**

ACLens is a PowerShell-based GUI tool that recursively reads NTFS permissions from any folder, generates interactive HTML reports, and lets you compare scan results over time to detect permission changes.

> ⚠️ **This is an early alpha release (v0.1.0).** Core functionality works but the tool is still in active development. Feedback and bug reports are very welcome.

---

## Screenshots

*Screenshots will be added in a future release.*

---

## Features

- 📂 **Recursive NTFS scan** — reads all subfolders with configurable depth limit
- 🔍 **Change detection** — highlights folders where permissions differ from the parent
- 📄 **Interactive HTML report** — filter, search, expand/collapse, print
- 📖 **Collapsible legend** — explains all permission types, scopes and inheritance options
- 💾 **JSON snapshot export** — automatic baseline for future comparisons
- 🔄 **Compare Scans** — diff two snapshots: shows added/removed folders and exact permission changes
- 🖥️ **Full GUI** — no command line knowledge required
- 🔒 **Admin check** — prompts to restart as Administrator for full coverage

---

## Requirements

| | |
|---|---|
| **OS** | Windows 10, Windows 11, Windows Server 2016, 2019, 2022, 2025 |
| **PowerShell** | 5.1 or higher (pre-installed on all supported versions) |
| **Permissions** | Read access to scanned folders. Administrator recommended. |

---

## Getting Started

### 1. Download

Download `ACLens.ps1` from the [Releases](../../releases) page.

### 2. Allow script execution

Windows blocks PowerShell scripts by default. Open PowerShell and run:

```powershell
# Option A — single run only (recommended for first try)
powershell.exe -ExecutionPolicy Bypass -File "C:\Path\To\ACLens.ps1"

# Option B — permanently allow for your user account
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### 3. Run ACLens

Right-click `ACLens.ps1` → **Run with PowerShell**, or use the command above.

> **Tip:** Run as Administrator for complete results on protected system folders.

---

## Usage

### Basic Scan

1. Click **Browse...** and select the folder to scan
2. Optionally set a custom output path and maximum scan depth
3. Click **Start Analysis**
4. The HTML report opens automatically in your browser

### Output Files

Both files are saved in the **same folder as `ACLens.ps1`** by default:

| File | Description |
|---|---|
| `ACLens_Report_YYYY-MM-DD_HH-mm-ss.html` | Interactive HTML permission report |
| `ACLens_Report_YYYY-MM-DD_HH-mm-ss.json` | JSON snapshot for baseline comparison |

### Compare Scans

1. Run a first scan to create a baseline JSON
2. Click **Compare Scans** in the footer
3. Select the baseline `.json` file
4. Enter the folder path for a new live scan
5. Click **Run Comparison** — the diff report is saved and opened automatically

The diff report shows:
- 🟢 **Added** — folders not present in the baseline
- 🔴 **Removed** — folders that no longer exist
- 🟡 **Changed** — exact permission differences per folder (added/removed rules, owner changes, inheritance changes)

---

## HTML Report

The report includes:

- **Header** — scanned path, date, computer, user, total folders
- **Statistics** — total / changed / inherited-only / errors
- **Legend** — collapsible reference for all status codes, permission types and scopes
- **Toolbar** — filter by status, full-text search, expand/collapse all, print
- **Folder list** — click any folder to expand and see the full permission table

### Folder Status

| Badge | Meaning |
|---|---|
| 🟣 Root | The start folder |
| 🟡 Changed | Permissions differ from parent — explicit rules set or inheritance modified |
| 🟢 Inherited only | All permissions passed down from parent, no local changes |
| 🔴 Error | Could not read permissions — access denied or protected system folder |

---

## Troubleshooting

| Problem | Solution |
|---|---|
| Script blocked on start | Use `-ExecutionPolicy Bypass` or set `RemoteSigned` (see Getting Started) |
| Folders show "Error" | Run as Administrator for full access to system folders |
| 0 folders in report | Verify the start path exists and is accessible |
| JSON baseline not found | Run a new scan first — JSON is saved automatically next to the HTML report |
| Browser doesn't open | Uncheck "Open report in browser" and open the HTML file manually |

---

## Known Limitations (v0.1.0-alpha)

- No dark/light theme toggle in the HTML report
- Compare Scans always uses unlimited depth for the live scan
- No export to CSV or Excel
- No scheduled/automated scan support built-in
- GUI is not DPI-aware on very high-resolution displays

---

## Roadmap

- [ ] Version tagging and automatic update check
- [ ] Scheduled scan support
- [ ] CSV / Excel export
- [ ] Side-by-side diff view in HTML
- [ ] DPI-aware GUI scaling
- [ ] Dark/light mode toggle in HTML report

---

## License

This project is licensed under the **MIT License** — see [LICENSE](LICENSE) for details.

---

## Contributing

This is an early alpha — bug reports, feature requests and pull requests are welcome!

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/my-feature`)
3. Commit your changes
4. Open a Pull Request

Please open an issue first for larger changes.

---

*Made with PowerShell + Windows Forms*
