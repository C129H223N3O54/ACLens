# ACLens

> **NTFS Permission Analyzer & Reporter for Windows**

ACLens is a PowerShell-based GUI tool that recursively reads NTFS permissions from any folder, generates interactive HTML reports, and lets you compare scan results over time to detect permission changes.

> ⚠️ **This is an early alpha release (v0.1.0).** Core functionality works but the tool is still in active development. Feedback and bug reports are very welcome.

---

## Features

- 📂 **Recursive NTFS scan** — reads all subfolders with configurable depth limit
- 🔍 **Change detection** — highlights folders where permissions differ from the parent
- 📄 **Interactive HTML report** — filter, search, expand/collapse, print
- 📖 **Collapsible legend** — explains all permission types, scopes and inheritance
- 💾 **JSON snapshot export** — automatic baseline saved alongside every HTML report
- 🔄 **Compare Scans** — diff two snapshots: added/removed folders and exact permission changes
- 🖥️ **Full GUI** — no command line knowledge required
- 🔒 **Admin check on startup** — prompts to restart as Administrator for full coverage
- 📁 **Output to script folder** — reports saved next to `ACLens.ps1` by default

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

Download `ACLens.ps1` from the [Releases](../../releases) page. No installation needed — single file.

### 2. Allow script execution

```powershell
# Option A — single run only (recommended)
powershell.exe -ExecutionPolicy Bypass -File "C:\Path\To\ACLens.ps1"

# Option B — permanently allow for your user account
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### 3. Run

Right-click `ACLens.ps1` → **Run with PowerShell**, or use the command above.

> ACLens will automatically prompt you to restart as Administrator if it detects it's running without elevated privileges.

---

## Usage

### Basic Scan

1. Click **Browse...** and select the folder to scan
2. Optionally set a custom output path and maximum depth (`0` = unlimited)
3. Click **Start Analysis**
4. The HTML report opens automatically in your browser

### Output Files

Saved in the **same folder as `ACLens.ps1`** by default:

| File | Description |
|---|---|
| `ACLens_Report_YYYY-MM-DD_HH-mm-ss.html` | Interactive HTML permission report |
| `ACLens_Report_YYYY-MM-DD_HH-mm-ss.json` | JSON snapshot for baseline comparison |

### Compare Scans

1. Run a first scan to create a baseline JSON snapshot
2. Click **Compare Scans** in the footer
3. Select the baseline `.json` file (auto-filled with last scan)
4. Enter the folder path for the new live scan
5. Click **Run Comparison** — diff report opens automatically

**Diff report shows:**
- 🟢 **Added** — folders not present in the baseline
- 🔴 **Removed** — folders that no longer exist
- 🟡 **Changed** — exact permission differences per folder

---

## HTML Report

| Section | Description |
|---|---|
| Header | Scanned path, date, computer, user, total folder count |
| Statistics | Total / Changed / Inherited-only / Errors |
| Legend | Collapsible reference for all status codes and permission types |
| Toolbar | Filter by status, full-text search, expand/collapse all, print |
| Folder list | Click any row to expand the full permission table |

### Folder Status

| Badge | Meaning |
|---|---|
| 🟣 Root | The start folder |
| 🟡 Changed | Permissions differ from parent |
| 🟢 Inherited only | All permissions passed down from parent, no local changes |
| 🔴 Error | Could not read permissions — access denied or system folder |

---

## Troubleshooting

| Problem | Solution |
|---|---|
| Script blocked | Use `-ExecutionPolicy Bypass` or set `RemoteSigned` |
| Folders show "Error" | Run as Administrator |
| 0 folders in report | Verify the start path exists and is accessible |
| JSON not found | Run a new scan first — JSON saves automatically next to the HTML |
| Browser doesn't open | Uncheck "Open report in browser" and open the file manually |

---

## Known Limitations (v0.1.0-alpha)

- Compare Scans always uses unlimited depth for the live scan
- No CSV / Excel export
- No scheduled scan support
- GUI not DPI-aware on very high-resolution displays

---

## Roadmap

- [ ] **v0.2.0** — SharePoint Online support (Graph API, Sites, Libraries, Groups, External Sharing)
- [ ] Scheduled scan support
- [ ] CSV / Excel export
- [ ] DPI-aware GUI scaling

---

## License

MIT — see [LICENSE](LICENSE) for details.

---

## Contributing

Bug reports, feature requests and pull requests are welcome!
Please read [CONTRIBUTING.md](CONTRIBUTING.md) first.

---

*Built with PowerShell 5.1 + Windows Forms*
