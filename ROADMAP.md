# ACLens — Roadmap

This document outlines the planned features and direction for future releases.
Priorities may shift based on feedback and real-world usage.

---

## ✅ Released

### v0.1.0-alpha
- Recursive NTFS permission scan
- Interactive HTML report (filter, search, expand/collapse)
- JSON snapshot export
- Compare Scans (NTFS diff)
- Windows Forms GUI
- Admin check on startup

### v0.2.0-alpha
- SharePoint Online tab
- Setup Wizard (Device Code Flow + Manual)
- SP Scanner (Sites, Libraries, Folders, Groups, External Sharing)
- SP HTML report
- SP Compare

### v0.2.1-alpha
- Bug fixes (PS 5.1 switch statement)
- Version badge in header
- Header layout improvements

---

## 🔜 v0.3.0 — Quality of Life

> Focus: making daily use smoother and reports more useful.

- [ ] **CSV / Excel export** — export permission lists directly to `.xlsx` for further analysis
- [ ] **Recently scanned paths** — dropdown with last 5 paths per tab
- [x] **Dark / Light mode toggle** in HTML report
- [ ] **Improved print / PDF output** — clean, paginated PDF directly from the tool
- [ ] **DPI-aware GUI** — correct scaling on high-resolution displays
- [ ] **Multi-select SharePoint sites** — scan multiple sites in one run

---

## 🔜 v0.4.0 — Automation & Alerting

> Focus: running ACLens without manual interaction.

- [ ] **Scheduled scans** — configure Windows Task Scheduler directly from the GUI
- [ ] **E-mail notifications** — send an alert when permissions change after a scheduled scan
- [ ] **Silent / headless mode** — run ACLens from the command line without opening the GUI (for scripting and automation)
- [ ] **Auto-compare on schedule** — automatically compare each new scan against the previous baseline

---

## 🔜 v0.5.0 — Intelligence & Risk

> Focus: turning raw permission data into actionable insights.

- [ ] **Risk Score** — automatic flagging of high-risk patterns:
  - `Everyone` or `Authenticated Users` with broad access
  - Excessive `Full Control` assignments
  - External users with write permissions
  - Inheritance disabled without explicit justification
- [ ] **Compliance Rules** — define custom rules (e.g. "no Everyone on production folders") with Pass / Fail results per folder
- [ ] **User-centric view** — "Where does User X have access?" — cross-folder search across all scanned paths
- [ ] **Nested group resolution** — show who *actually* has access through nested AD/AAD groups

---

## 🔜 v0.6.0 — Extended Sources

> Focus: covering more platforms beyond NTFS and SharePoint.

- [ ] **OneDrive for Business** — personal and shared drives via Graph API
- [ ] **Microsoft Teams** — channel and tab permission scanning
- [ ] **Azure Files** — SMB shares in Azure Storage
- [ ] **Exchange / Mailbox permissions** — shared mailboxes, calendar permissions

---

## 💡 Future Ideas (Backlog)

These are not yet scheduled but worth exploring:

- **Timeline view** — visualize permission changes across multiple snapshots over time
- **Azure AD Deep Dive** — show AAD group membership, license, last sign-in per principal
- **Web Dashboard** — ACLens as a lightweight local web server with a live browser interface
- **Active Directory integration** — resolve on-premise AD groups and OU structure
- **Side-by-side diff view** in HTML compare report
- **Report branding** — custom logo and company name in HTML reports

---

## 💬 Feedback & Suggestions

Have an idea that's not on this list?
[Open a feature request](../../issues/new?labels=enhancement) — all suggestions are welcome.

---

*This roadmap reflects current thinking and is subject to change.*
