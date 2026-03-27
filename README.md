# ACLens

> **NTFS & SharePoint Permission Analyzer & Reporter for Windows**

ACLens is a PowerShell-based GUI tool that reads NTFS and SharePoint Online permissions, generates interactive HTML reports, and lets you compare scan results over time to detect permission changes.

> ⚠️ **This is an early alpha release (v0.2.1).** Core functionality works but the tool is still in active development. Feedback and bug reports are very welcome.

---

## Origin & Credits

The idea for ACLens came from **Philipp Herrmann** — born from a real-world need. Without that spark, this tool wouldn't exist. 🙌

The idea was picked up and driven by **Jan Erik Mueller**, who defined the requirements, direction and scope of the project.

The entire codebase — every line of PowerShell, the Windows Forms GUI, the scan engine, all dialogs and tools — was written by **[Claude](https://claude.ai)**, an AI assistant made by **[Anthropic](https://anthropic.com)**.

This project is an example of human–AI collaboration: people with a vision, and an AI that implements it.

---

## Features

### NTFS
- 📂 **Recursive scan** with configurable depth limit
- 🔍 **Change detection** — highlights folders where permissions differ from the parent
- 📄 **Interactive HTML report** — filter, search, expand/collapse, print
- 📖 **Collapsible legend** — explains all permission types, scopes and inheritance
- 💾 **JSON snapshot** — auto-saved alongside every HTML report for comparison
- 🔄 **Compare Scans** — diff two snapshots: added/removed folders + exact permission changes

### SharePoint Online
- ☁️ **Full site scan** — Sites, Document Libraries, Folders, Groups, External Sharing
- 🔐 **Two auth modes** — Automatic (Device Code, zero manual steps) or Manual (Client ID/Secret)
- 📄 **SP HTML report** — Site permissions, Groups & Members, Library/Folder tree with unique permission detection
- 🔄 **SP Compare** — diff two SP snapshots across time

### General
- 🖥️ **Full GUI** — no command line knowledge required
- 🔒 **Admin check** — prompts to restart as Administrator on startup
- 📁 **Output to script folder** — all reports saved next to `ACLens.ps1`

---

## Requirements

| | |
|---|---|
| **OS** | Windows 10, Windows 11, Windows Server 2016–2025 |
| **PowerShell** | 5.1 or higher (pre-installed on all supported versions) |
| **For NTFS** | Read access to scanned folders. Administrator recommended. |
| **For SharePoint** | Microsoft 365 account with Application Administrator role (first-time setup only) |

---

## Getting Started

### 1. Download

Download `ACLens.ps1` from the [Releases](../../releases) page. Single file — no installation needed.

### 2. Allow script execution

```powershell
# Option A — single run (recommended for first try)
powershell.exe -ExecutionPolicy Bypass -File "C:\Path\To\ACLens.ps1"

# Option B — permanently allow for your user account
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### 3. Run

Right-click `ACLens.ps1` → **Run with PowerShell**, or use Option A.

> ACLens will prompt you to restart as Administrator if it detects missing privileges.

---

## NTFS Usage

1. Click the **NTFS** tab
2. Click **Browse...** and select the folder to scan
3. Optionally set a custom output path and maximum depth (`0` = unlimited)
4. Click **Start Analysis**
5. The HTML report opens automatically in your browser

**Output files** (saved next to `ACLens.ps1`):

| File | Description |
|---|---|
| `ACLens_Report_YYYY-MM-DD_HH-mm-ss.html` | Interactive HTML report |
| `ACLens_Report_YYYY-MM-DD_HH-mm-ss.json` | Snapshot for comparison |

---

## SharePoint Usage

### First-time Setup

1. Click the **SharePoint** tab
2. Click **Setup / Reconnect**
3. **Automatic (recommended):** Click **Get Device Code**, open `microsoft.com/devicelogin`, sign in with your admin account — ACLens creates the App Registration automatically
4. **Manual:** Enter Tenant ID, Client ID and Client Secret — see [Manual Setup Guide](docs/manual-sp-setup.md)

### Running a Scan

1. Enter the SharePoint Site URL
2. Set maximum depth (`0` = unlimited)
3. Click **Start Analysis**

**Output files:**

| File | Description |
|---|---|
| `ACLens_SP_Report_YYYY-MM-DD_HH-mm-ss.html` | Interactive SP HTML report |
| `ACLens_SP_Report_YYYY-MM-DD_HH-mm-ss.json` | SP snapshot for comparison |

---

## Compare Scans

Click **Compare Scans** in the footer to open the comparison window.

**NTFS Compare:**
1. Select the baseline `.json` file (auto-filled with last scan)
2. Enter the folder path for a new live scan
3. Click **Run Comparison**

**SharePoint Compare:**
1. Select the baseline SP `.json` file
2. Enter the SharePoint Site URL for the new live scan
3. Click **Run Comparison**

**Both diff reports show:**
- 🟢 **Added** — new folders not in the baseline
- 🔴 **Removed** — folders that no longer exist
- 🟡 **Changed** — exact permission differences (added/removed rules, ownership, inheritance)

---

## HTML Report Reference

### NTFS Report

| Section | Description |
|---|---|
| Header | Path, date, computer, user, total folders |
| Statistics | Total / Changed / Inherited-only / Errors |
| Legend | Collapsible reference (click to expand) |
| Toolbar | Filter, search, expand/collapse all, print |
| Folder list | Click any row to see full permission table |

### Folder Status

| Badge | Meaning |
|---|---|
| 🟣 Root | Start folder |
| 🟡 Changed | Permissions differ from parent |
| 🟢 Inherited only | All permissions passed down — no local changes |
| 🔴 Error | Could not read — access denied or system folder |

### SharePoint Report

| Section | Description |
|---|---|
| Header | Site URL, name, date, statistics |
| Site Permissions | Direct site-level permission assignments |
| Groups & Members | All SharePoint groups with member lists |
| Libraries | Expandable per library with folder permission tree |

---

## Troubleshooting

| Problem | Solution |
|---|---|
| Script blocked | Use `-ExecutionPolicy Bypass` or set `RemoteSigned` |
| NTFS folders show "Error" | Run as Administrator |
| 0 folders in report | Verify the path exists and is accessible |
| SP: "Not configured" | Use Setup / Reconnect in the SharePoint tab |
| SP: "Insufficient privileges" | Re-grant admin consent in Azure Portal |
| SP: Site not found | Copy the URL exactly from the SharePoint browser address bar |
| JSON not found for compare | Run a new scan first — JSON saves automatically |

---

## Known Limitations (v0.2.1-alpha)

- SP Compare runs a full live scan (depth limit not applied)
- No CSV / Excel export yet
- No scheduled scan support
- GUI not DPI-aware on very high-resolution displays

---

## Roadmap

- [ ] **v0.3.0** — Scheduled scans, CSV/Excel export, DPI-aware GUI
- [ ] Side-by-side diff view in HTML
- [ ] Multi-site SP scanning

---

## Documentation

- [Manual SharePoint Setup](manual-sp-setup.md) — step-by-step Azure App Registration guide

---

## License

MIT — see [LICENSE](LICENSE) for details.

---

## Contributing

Bug reports, feature requests and pull requests are welcome!
Please read [CONTRIBUTING.md](CONTRIBUTING.md) first.

---

*Built with PowerShell 5.1 + Windows Forms*
