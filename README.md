# ACLens

> **NTFS & SharePoint Permission Analyzer & Reporter for Windows**

ACLens is a PowerShell-based GUI tool that reads NTFS and SharePoint Online permissions, generates interactive HTML reports, and lets you compare scan results over time to detect permission changes.

> ⚠️ **This is a beta release (v0.4.0).** Core functionality works but the tool is still in active development. Feedback and bug reports are very welcome.

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
- 📄 **Interactive HTML report** — paginated, filter, search, expand/collapse, print
- 📖 **Collapsible legend** — explains all permission types, scopes and inheritance
- 💾 **JSON snapshot** — auto-saved alongside every HTML report
- 🔄 **Compare Scans** — diff two snapshots: added/removed folders + exact permission changes

### SharePoint Online
- ☁️ **Full site scan** — Sites, Libraries, Folders, Groups, External Sharing
- 🔐 **Two auth modes** — Automatic (Device Code) or Manual (Client ID/Secret)
- 📄 **SP HTML report** — site permissions, groups, library/folder tree
- 🔄 **SP Compare** — diff two SP snapshots

### General
- 🖥️ **Full GUI** — no command line required
- 🌍 **DE/EN language toggle** — switch between German and English at any time
- 🌙 **Dark/Light mode** — for both GUI and HTML reports
- 🔒 **Admin check** on startup
- 📁 **Output to script folder** by default

---

## Requirements

| | |
|---|---|
| **OS** | Windows 10, Windows 11, Windows Server 2016–2025 |
| **PowerShell** | 5.1 or higher |
| **For NTFS** | Read access to scanned folders. Administrator recommended. |
| **For SharePoint** | Microsoft 365 account with Application Administrator role (first-time setup only) |

---

## Getting Started

### 1. Download

Download `ACLens.ps1` from the [Releases](../../releases) page. Single file — no installation needed.

### 2. Allow script execution

```powershell
# Option A — single run
powershell.exe -ExecutionPolicy Bypass -File "C:\Path\To\ACLens.ps1"

# Option B — permanently for your user
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### 3. Run

Right-click `ACLens.ps1` → **Run with PowerShell**.

---

## NTFS Usage

1. Select the **NTFS** tab
2. Click **Browse...** and select a folder
3. Set maximum depth (`0` = unlimited)
4. Click **Start Analysis**

**Output files** saved next to `ACLens.ps1`:

| File | Description |
|---|---|
| `ACLens_Report_YYYY-MM-DD_HH-mm-ss.html` | Interactive HTML report |
| `ACLens_Report_YYYY-MM-DD_HH-mm-ss.json` | Snapshot for comparison |

---

## SharePoint Usage

### First-time Setup

1. Click the **SharePoint** tab
2. Click **Setup / Reconnect**
3. **Automatic:** Click **Get Device Code**, open `microsoft.com/devicelogin`, sign in — ACLens handles the rest
4. **Manual:** Enter Tenant ID, Client ID and Client Secret — see [Manual Setup Guide (EN)](manual-sp-setup.md) / [Manuelle Einrichtung (DE)](manuelle-sp-einrichtung.md)

---

## Compare Scans

Click **Compare Scans** in the footer to diff two scan results.

Both **NTFS** and **SharePoint** comparisons are supported.

| Status | Meaning |
|---|---|
| 🟢 Added | New folders not in baseline |
| 🔴 Removed | Folders that no longer exist |
| 🟡 Changed | Exact permission differences |

---

## Troubleshooting

| Problem | Solution |
|---|---|
| Script blocked | Use `-ExecutionPolicy Bypass` or set `RemoteSigned` |
| Folders show "Error" | Run as Administrator |
| SP: "Not configured" | Use Setup / Reconnect |
| SP: "Insufficient privileges" | Re-grant admin consent in Azure Portal |
| JSON not found | Run a new scan first |

---

## Documentation

- [Manual SharePoint Setup (EN)](manual-sp-setup.md)
- [Manuelle SharePoint Einrichtung (DE)](manuelle-sp-einrichtung.md)
- [Roadmap](ROADMAP.md)

---

## License

MIT — see [LICENSE](LICENSE) for details.

---

## Contributing

Bug reports and pull requests welcome — see [CONTRIBUTING.md](CONTRIBUTING.md).

---

*Built with PowerShell 5.1 + Windows Forms*
