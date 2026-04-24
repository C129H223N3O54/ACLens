#Requires -Version 5.1
<#
.SYNOPSIS
    ACLens - NTFS Permission Analyzer & Reporter

.DESCRIPTION
    Entry point. Loads all modules and launches the GUI.

.NOTES
    Version:    0.4.0-beta
    Author:     ACLens Contributors
    Compatible: Windows 10/11, Server 2016-2025, PowerShell 5.1+
#>

# ── Admin check ──────────────────────────────────────────────
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    function HAC([string]$h) {
        $h = $h.TrimStart('#')
        [System.Drawing.Color]::FromArgb(
            [Convert]::ToInt32($h.Substring(0,2),16),
            [Convert]::ToInt32($h.Substring(2,2),16),
            [Convert]::ToInt32($h.Substring(4,2),16))
    }

    $adm = [System.Windows.Forms.Form]::new()
    $adm.Text            = "ACLens"
    $adm.ClientSize      = [System.Drawing.Size]::new(500, 267)
    $adm.StartPosition   = "CenterScreen"
    $adm.FormBorderStyle = "FixedDialog"
    $adm.MaximizeBox     = $false
    $adm.MinimizeBox     = $false
    $adm.BackColor       = HAC '#1E1E2E'
    $adm.Font            = [System.Drawing.Font]::new("Segoe UI", 9)

    $aHdr = [System.Windows.Forms.Panel]::new()
    $aHdr.SetBounds(0, 0, 500, 56)
    $aHdr.BackColor = HAC '#1A1A2E'
    $adm.Controls.Add($aHdr)

    $aTitle = [System.Windows.Forms.Label]::new()
    $aTitle.Text = "ACLens"; $aTitle.Font = [System.Drawing.Font]::new("Segoe UI", 15, [System.Drawing.FontStyle]::Bold)
    $aTitle.ForeColor = HAC '#A78BFA'; $aTitle.BackColor = [System.Drawing.Color]::Transparent
    $aTitle.AutoSize = $true; $aTitle.Location = [System.Drawing.Point]::new(18, 13)
    $aHdr.Controls.Add($aTitle)

    $aSub = [System.Windows.Forms.Label]::new()
    $aSub.Text = "Administrator Recommended"; $aSub.Font = [System.Drawing.Font]::new("Segoe UI", 9)
    $aSub.ForeColor = HAC '#6B7280'; $aSub.BackColor = [System.Drawing.Color]::Transparent
    $aSub.AutoSize = $true; $aSub.Location = [System.Drawing.Point]::new(130, 21)
    $aHdr.Controls.Add($aSub)

    $aHdrSep = [System.Windows.Forms.Panel]::new()
    $aHdrSep.SetBounds(0, 56, 500, 1); $aHdrSep.BackColor = HAC '#4B5563'
    $adm.Controls.Add($aHdrSep)

    $aIcon = [System.Windows.Forms.Label]::new()
    $aIcon.Text = [char]0x26A0; $aIcon.Font = [System.Drawing.Font]::new("Segoe UI", 28)
    $aIcon.ForeColor = HAC '#FBBF24'; $aIcon.BackColor = [System.Drawing.Color]::Transparent
    $aIcon.AutoSize = $true; $aIcon.Location = [System.Drawing.Point]::new(18, 66)
    $adm.Controls.Add($aIcon)

    $aMsg1 = [System.Windows.Forms.Label]::new()
    $aMsg1.Text = "ACLens is not running as Administrator."
    $aMsg1.Font = [System.Drawing.Font]::new("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $aMsg1.ForeColor = HAC '#E2E8F0'; $aMsg1.BackColor = [System.Drawing.Color]::Transparent
    $aMsg1.AutoSize = $true; $aMsg1.Location = [System.Drawing.Point]::new(68, 68)
    $adm.Controls.Add($aMsg1)

    $aMsg2 = [System.Windows.Forms.Label]::new()
    $aMsg2.Text = "Without administrator privileges some folders may not be readable and permission results could be incomplete."
    $aMsg2.Font = [System.Drawing.Font]::new("Segoe UI", 9)
    $aMsg2.ForeColor = HAC '#9CA3AF'; $aMsg2.BackColor = [System.Drawing.Color]::Transparent
    $aMsg2.Size = [System.Drawing.Size]::new(414, 42); $aMsg2.Location = [System.Drawing.Point]::new(68, 96)
    $adm.Controls.Add($aMsg2)

    $aMsg3 = [System.Windows.Forms.Label]::new()
    $aMsg3.Text = "Restart ACLens as Administrator now?"
    $aMsg3.Font = [System.Drawing.Font]::new("Segoe UI", 9)
    $aMsg3.ForeColor = HAC '#E2E8F0'; $aMsg3.BackColor = [System.Drawing.Color]::Transparent
    $aMsg3.AutoSize = $true; $aMsg3.Location = [System.Drawing.Point]::new(68, 148)
    $adm.Controls.Add($aMsg3)

    $aSep = [System.Windows.Forms.Panel]::new()
    $aSep.SetBounds(0, 196, 500, 1); $aSep.BackColor = HAC '#4B5563'
    $adm.Controls.Add($aSep)

    $aYes = [System.Windows.Forms.Button]::new()
    $aYes.Text = "Yes, restart as Admin"; $aYes.Size = [System.Drawing.Size]::new(180, 38)
    $aYes.Location = [System.Drawing.Point]::new(18, 214); $aYes.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $aYes.Font = [System.Drawing.Font]::new("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $aYes.BackColor = HAC '#7C3AED'; $aYes.ForeColor = [System.Drawing.Color]::White
    $aYes.Cursor = "Hand"; $aYes.FlatAppearance.BorderColor = HAC '#4C1D95'
    $aYes.FlatAppearance.BorderSize = 1
    $aYes.DialogResult = [System.Windows.Forms.DialogResult]::Yes
    $adm.Controls.Add($aYes)

    $aNo = [System.Windows.Forms.Button]::new()
    $aNo.Text = "Continue without Admin"; $aNo.Size = [System.Drawing.Size]::new(180, 38)
    $aNo.Location = [System.Drawing.Point]::new(208, 214); $aNo.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $aNo.Font = [System.Drawing.Font]::new("Segoe UI", 9)
    $aNo.BackColor = HAC '#374151'; $aNo.ForeColor = HAC '#E2E8F0'
    $aNo.Cursor = "Hand"; $aNo.FlatAppearance.BorderColor = HAC '#4B5563'
    $aNo.FlatAppearance.BorderSize = 1
    $aNo.DialogResult = [System.Windows.Forms.DialogResult]::No
    $adm.Controls.Add($aNo)

    $adm.AcceptButton = $aYes
    $adm.CancelButton = $aNo

    if ($adm.ShowDialog() -eq [System.Windows.Forms.DialogResult]::Yes) {
        Start-Process -FilePath "powershell.exe" `
            -ArgumentList "-ExecutionPolicy Bypass -File `"$PSCommandPath`"" `
            -Verb RunAs
        exit
    }
}

# ── Load assemblies ───────────────────────────────────────────

# ============================================================
# SPRACHE / LANGUAGE
# ============================================================
$script:lang = "EN"

$script:STR = @{}

function Set-Language($l) {
    $script:lang = $l
    if ($l -eq "DE") {
        $script:STR = @{
            # Header
            AppSubtitle     = "NTFS & SharePoint Berechtigungsanalyse"
            # Tabs
            TabNTFS         = "  NTFS"
            TabSP           = "  SharePoint"
            # NTFS fields
            LblPath         = "Startpfad:"
            LblOutput       = "Ausgabedatei (HTML + JSON):"
            LblDepth        = "Maximale Tiefe:"
            LblUnlimited    = "  0 = unbegrenzt"
            LblBrowse       = "Durchsuchen..."
            LblSaveAs       = "Speichern..."
            ChkBrowser      = "Bericht nach Erstellung im Browser öffnen"
            LblReady        = "Bereit."
            # SP fields
            SpNotConfigured = "Nicht konfiguriert"
            SpConnected     = "Verbunden"
            SpConnectHint   = "Klicke 'Einrichten / Verbinden' um ACLens mit deinem Microsoft 365 Mandanten zu verbinden."
            SpLblUrl        = "SharePoint-Website URL:"
            SpLblDepth      = "Maximale Tiefe:"
            SpChkBrowser    = "Bericht nach Erstellung im Browser öffnen"
            SpBtnSetup      = "Einrichten / Verbinden"
            # Footer buttons
            BtnStart        = "Starten"
            BtnCancel       = "Abbrechen"
            BtnOpenLast     = "Letzter Bericht"
            BtnChangelog    = "Changelog"
            BtnCompare      = "Vergleichen"
            BtnThemeLight   = "Hell"
            BtnThemeDark    = "Dunkel"
            BtnLang         = "DE"
            # Status messages
            StepScanFolders = "Schritt 1/3: Ordner werden gescannt..."
            StepReadPerms   = "Schritt 2/3: NTFS-Berechtigungen werden gelesen"
            StepGenReport   = "Schritt 3/3: HTML-Bericht wird erstellt..."
            StepCancelled   = "Abgebrochen."
            StepDone        = "Fertig! Bericht gespeichert:"
            StepError       = "Fehler:"
            MsgInvalidPath  = "Bitte einen gültigen Ordnerpfad angeben."
            MsgInvalidPathT = "Ungültiger Pfad"
            # Compare window
            CmpTitle        = "ACLens - Scans vergleichen"
            CmpNTFS         = "  NTFS-Vergleich"
            CmpSP           = "  SharePoint-Vergleich"
            CmpBaseline     = "Baseline (JSON):"
            CmpScanPath     = "Aktueller Scanpfad:"
            CmpSpBaseline   = "SharePoint Baseline-Snapshot (JSON):"
            CmpSpUrl        = "Aktuelle SharePoint-Website URL:"
            CmpSpNote       = "Hinweis: Es wird ein neuer SP-Scan durchgeführt und mit der Baseline verglichen."
            CmpRun          = "Vergleich starten"
            CmpReady        = "Bereit."
            CmpLoading      = "Baseline wird geladen..."
            CmpScanning     = "Schritt 2/2: Ordner werden gescannt..."
            CmpGenerating   = "Diff-Bericht wird erstellt..."
            CmpDone         = "Fertig!"
            CmpMissingJson  = "Bitte eine gültige Baseline-JSON-Datei auswählen."
            CmpMissingPath  = "Bitte einen gültigen Scanpfad angeben."
            CmpMissingUrl   = "Bitte eine gültige SharePoint-Website URL angeben."
            CmpMissingInput = "Fehlende Eingabe"
            # SP Setup Wizard
            WizTitle        = "ACLens — SharePoint Einrichtung"
            WizSub          = "Mit Microsoft 365 verbinden"
            WizAutoTab      = "Automatische Einrichtung"
            WizManualTab    = "Manuelle Einrichtung"
            WizAutoInfo     = "ACLens erstellt automatisch eine App-Registrierung in deinem Azure AD Mandanten über den Gerätecode-Flow. Du benötigst ein Konto mit der Rolle 'Anwendungsadministrator' oder 'Globaler Administrator'."
            WizStep1        = "Schritt 1 — Klicke auf 'Gerätecode anfordern'"
            WizStep2        = "Schritt 2 — Öffne den Link und gib den unten angezeigten Code ein"
            WizStep3        = "Schritt 3 — Melde dich mit deinem Administratorkonto an — ACLens erledigt den Rest"
            WizGetCode      = "Gerätecode anfordern"
            WizOpenBrowser  = "microsoft.com/devicelogin öffnen"
            WizWaiting      = "Warte auf Anmeldung..."
            WizCreating     = "Angemeldet! App-Registrierung wird erstellt..."
            WizDone         = "App-Registrierung erstellt!"
            WizCancelled    = "Anmeldung abgebrochen oder abgelaufen. Bitte erneut versuchen."
            WizManualInfo   = "Zugangsdaten einer vorhandenen Azure AD App-Registrierung manuell eingeben. Siehe Dokumentation für Einrichtungsanweisungen."
            WizOpenDocs     = "Anleitung zur manuellen Einrichtung öffnen"
            WizTenantId     = "Mandanten-ID"
            WizClientId     = "Client-ID"
            WizSecret       = "Geheimer Clientschlüssel"
            WizSaveConnect  = "Speichern && verbinden"
            WizMissingData  = "Bitte alle Felder ausfüllen."
            WizSaved        = "Zugangsdaten erfolgreich gespeichert."
            WizClose        = "Schließen"
            # HTML Report
            RptTitle        = "ACLens — NTFS-Berechtigungsbericht"
            RptStatTotal    = "Ordner gesamt"
            RptStatChanged  = "Geändert"
            RptStatInherited= "Nur geerbt"
            RptStatErrors   = "Fehler"
            RptLegendBtn    = "📖 Legende & Referenz"
            RptStatusRoot   = "Stamm"
            RptStatusChanged= "Geändert"
            RptStatusInherited = "Geerbt"
            RptStatusError  = "Fehler"
            RptFolderStatus = "Ordnerstatus"
            RptPermTypes    = "Berechtigungstypen"
            RptRuleTypes    = "Regeltypen"
            RptFCDesc       = "Lesen, Schreiben, Berechtigungen ändern, Besitz übernehmen"
            RptModifyDesc   = "Lesen, Schreiben, Dateien und Unterordner löschen"
            RptRXDesc       = "Dateien anzeigen und ausführen"
            RptReadDesc     = "Dateien und Ordnerinhalte anzeigen"
            RptWriteDesc    = "Dateien und Unterordner erstellen"
            RptExplicit     = "Explizit"
            RptInherited    = "Geerbt"
            RptAllow        = "Erlauben"
            RptDeny         = "Verweigern"
            RptExplicitSet  = "Direkt auf diesem Ordner gesetzt"
            RptInheritedFrom= "Vom übergeordneten Ordner weitergegeben"
            RptAllowDesc    = "Gewährt Zugriff — kann durch Verweigern überschrieben werden"
            RptDenyDesc     = "Sperrt Zugriff — überschreibt immer Erlauben"
            RptFilterAll    = "Alle"
            RptFilterChanged= "Geändert"
            RptFilterInh    = "Nur geerbt"
            RptFilterErr    = "Fehler"
            RptExpandPage   = "Seite aufklappen"
            RptCollapsePage = "Seite zuklappen"
            RptPrint        = "🖨 Drucken"
            RptPrincipal    = "Konto / Gruppe"
            RptType         = "Typ"
            RptPermission   = "Berechtigung"
            RptAppliesTo    = "Gilt für"
            RptKind         = "Art"
            RptSameAsParent = "✓ Wie übergeordneter Ordner — keine lokalen Änderungen"
            RptNoRules      = "Keine Berechtigungsregeln gefunden."
            RptFooter       = "ACLens v0.4.0-beta"
            RptThemeLight   = "☀️ Hell"
            RptThemeDark    = "🌙 Dunkel"
            RptPerPage      = "pro Seite"
            RptNoResults    = "Keine Ordner entsprechen diesem Filter."
            RptOwner        = "Besitzer"
            RptInheritance  = "Vererbung"
            RptExplicitPerms= "Explizite Berechtigungen"
            RptInheritedPerms="Geerbte Berechtigungen"
        }
    } else {
        $script:STR = @{
            AppSubtitle     = "NTFS & SharePoint Permission Analyzer"
            TabNTFS         = "  NTFS"
            TabSP           = "  SharePoint"
            LblPath         = "Start Path:"
            LblOutput       = "Output File (HTML + JSON):"
            LblDepth        = "Maximum Depth:"
            LblUnlimited    = "  0 = unlimited"
            LblBrowse       = "Browse..."
            LblSaveAs       = "Save as..."
            ChkBrowser      = "Open report in browser after creation"
            LblReady        = "Ready."
            SpNotConfigured = "Not configured"
            SpConnected     = "Connected"
            SpConnectHint   = "Click 'Setup / Reconnect' to register ACLens with your Microsoft 365 tenant."
            SpLblUrl        = "SharePoint Site URL:"
            SpLblDepth      = "Maximum Depth:"
            SpChkBrowser    = "Open report in browser after creation"
            SpBtnSetup      = "Setup / Reconnect"
            BtnStart        = "Start Analysis"
            BtnCancel       = "Cancel"
            BtnOpenLast     = "Open Last Report"
            BtnChangelog    = "Changelog"
            BtnCompare      = "Compare Scans"
            BtnThemeLight   = "Light"
            BtnThemeDark    = "Dark"
            BtnLang         = "EN"
            StepScanFolders = "Step 1/3: Scanning folders..."
            StepReadPerms   = "Step 2/3: Reading NTFS permissions"
            StepGenReport   = "Step 3/3: Generating HTML report..."
            StepCancelled   = "Cancelled."
            StepDone        = "Done!  Report saved to:"
            StepError       = "Error:"
            MsgInvalidPath  = "Please enter a valid folder path."
            MsgInvalidPathT = "Invalid Path"
            CmpTitle        = "ACLens - Compare Scans"
            CmpNTFS         = "  NTFS Compare"
            CmpSP           = "  SharePoint Compare"
            CmpBaseline     = "Baseline (JSON):"
            CmpScanPath     = "Current Scan Path:"
            CmpSpBaseline   = "Baseline SP Snapshot (JSON):"
            CmpSpUrl        = "Current SharePoint Site URL:"
            CmpSpNote       = "Note: A new live SP scan will run and be compared against the baseline JSON."
            CmpRun          = "Run Comparison"
            CmpReady        = "Ready."
            CmpLoading      = "Loading baseline..."
            CmpScanning     = "Step 2/2: Scanning folders..."
            CmpGenerating   = "Generating diff report..."
            CmpDone         = "Done!"
            CmpMissingJson  = "Please select a valid baseline JSON file."
            CmpMissingPath  = "Please enter a valid scan path."
            CmpMissingUrl   = "Please enter a valid SharePoint Site URL."
            CmpMissingInput = "Missing input"
            WizTitle        = "ACLens — SharePoint Setup"
            WizSub          = "Connect ACLens to Microsoft 365"
            WizAutoTab      = "Automatic Setup"
            WizManualTab    = "Manual Setup"
            WizAutoInfo     = "ACLens will automatically create an App Registration in your Azure AD tenant using the Device Code flow. You need an account with Application Administrator or Global Administrator role."
            WizStep1        = "Step 1 — Click 'Get Device Code' below"
            WizStep2        = "Step 2 — Open the link and enter the code shown below"
            WizStep3        = "Step 3 — Sign in with your admin account — ACLens handles the rest"
            WizGetCode      = "Get Device Code"
            WizOpenBrowser  = "Open microsoft.com/devicelogin"
            WizWaiting      = "Waiting for sign-in..."
            WizCreating     = "Signed in! Creating App Registration..."
            WizDone         = "App Registration created!"
            WizCancelled    = "Sign-in cancelled or expired. Try again."
            WizManualInfo   = "Manually enter credentials from an existing Azure AD App Registration. See the documentation for setup instructions."
            WizOpenDocs     = "Open Manual Setup Guide"
            WizTenantId     = "Tenant ID"
            WizClientId     = "Client ID"
            WizSecret       = "Client Secret"
            WizSaveConnect  = "Save && Connect"
            WizMissingData  = "Please fill in all fields."
            WizSaved        = "Credentials saved successfully."
            WizClose        = "Close"
            RptTitle        = "ACLens — NTFS Permission Report"
            RptStatTotal    = "Total folders"
            RptStatChanged  = "Changed"
            RptStatInherited= "Inherited only"
            RptStatErrors   = "Errors"
            RptLegendBtn    = "📖 Legend & Reference"
            RptStatusRoot   = "Root"
            RptStatusChanged= "Changed"
            RptStatusInherited = "Inherited"
            RptStatusError  = "Error"
            RptFolderStatus = "Folder Status"
            RptPermTypes    = "Permission Types"
            RptRuleTypes    = "Rule Types"
            RptFCDesc       = "Read, write, change permissions, take ownership"
            RptModifyDesc   = "Read, write, delete files and subfolders"
            RptRXDesc       = "View and run files"
            RptReadDesc     = "View files and folder contents"
            RptWriteDesc    = "Create files and subfolders"
            RptExplicit     = "Explicit"
            RptInherited    = "Inherited"
            RptAllow        = "Allow"
            RptDeny         = "Deny"
            RptExplicitSet  = "Set directly on this folder"
            RptInheritedFrom= "Passed down from a parent folder"
            RptAllowDesc    = "Grants access — can be overridden by Deny"
            RptDenyDesc     = "Blocks access — always overrides Allow"
            RptFilterAll    = "All"
            RptFilterChanged= "Changed"
            RptFilterInh    = "Inherited only"
            RptFilterErr    = "Errors"
            RptExpandPage   = "Expand page"
            RptCollapsePage = "Collapse page"
            RptPrint        = "🖨 Print"
            RptPrincipal    = "Principal"
            RptType         = "Type"
            RptPermission   = "Permission"
            RptAppliesTo    = "Applies to"
            RptKind         = "Kind"
            RptSameAsParent = "✓ Same as parent folder — no local changes"
            RptNoRules      = "No permission rules found."
            RptFooter       = "ACLens v0.4.0-beta"
            RptThemeLight   = "☀️ Light"
            RptThemeDark    = "🌙 Dark"
            RptPerPage      = "/ page"
            RptNoResults    = "No folders match this filter."
            RptOwner        = "Owner"
            RptInheritance  = "Inheritance"
            RptExplicitPerms= "Explicit Permissions"
            RptInheritedPerms="Inherited Permissions"
        }
    }
}

# Initialize with English
Set-Language "EN"


Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Web

# ── Load modules ─────────────────────────────────────────────
$moduleRoot = Join-Path $PSScriptRoot "modules"

# ============================================================
# MODULE: Core
# ============================================================
# ACLens - Core.psm1
# Shared colors, fonts, and helper functions used by all modules

function HC([string]$h) {
    $h = $h.TrimStart('#')
    [System.Drawing.Color]::FromArgb(
        [Convert]::ToInt32($h.Substring(0,2),16),
        [Convert]::ToInt32($h.Substring(2,2),16),
        [Convert]::ToInt32($h.Substring(4,2),16))
}

# ── Color palette (DiskLens standard) ────────────────────────
$script:COL = @{
    BgMain      = HC '#1E1E2E'
    BgInput     = HC '#2D2D3F'
    BgTitle     = HC '#1A1A2E'
    BgStatus    = HC '#111827'
    BgMid       = HC '#252535'
    BtnPrimary  = HC '#7C3AED'
    BtnPrimBrd  = HC '#4C1D95'
    BtnNeutral  = HC '#374151'
    BtnCancel   = HC '#991B1B'
    BtnCnclBrd  = HC '#7F1D1D'
    BtnBlue     = HC '#1D4ED8'
    BtnBlueBrd  = HC '#1E40AF'
    Border      = HC '#4B5563'
    BorderDark  = HC '#374151'
    TxtHi       = HC '#E2E8F0'
    TxtAccent   = HC '#A78BFA'
    TxtMid      = HC '#9CA3AF'
    TxtLow      = HC '#6B7280'
    TxtOk       = HC '#34D399'
    TxtErr      = HC '#F87171'
    TxtWarn     = HC '#FBBF24'
    BgSel       = HC '#4C1D95'
}

# ── Fonts ─────────────────────────────────────────────────────
$script:FN  = [System.Drawing.Font]::new("Segoe UI",  9)
$script:FB  = [System.Drawing.Font]::new("Segoe UI",  9,  [System.Drawing.FontStyle]::Bold)
$script:FT  = [System.Drawing.Font]::new("Segoe UI", 15, [System.Drawing.FontStyle]::Bold)
$script:FS  = [System.Drawing.Font]::new("Segoe UI",  8)
$script:FBTN = [System.Drawing.Font]::new("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)

# ── Layout constants ─────────────────────────────────────────
$script:PAD   = 20
$script:HDR_H = 54
$script:FTR_H = 72
$script:ROW_H = 54
$script:BTN_W = 110
$script:BTN_H = 26


# ============================================================
# MODULE: NTFS
# ============================================================
# ACLens - NTFS.psm1
# NTFS permission scanning and comparison logic

function Get-InheritanceFlagsDescription {
    param(
        [System.Security.AccessControl.InheritanceFlags]$InheritanceFlag,
        [System.Security.AccessControl.PropagationFlags]$PropagationFlag
    )
    $thisFolder  = $true
    $subFolders  = ($InheritanceFlag -band [System.Security.AccessControl.InheritanceFlags]::ContainerInherit) -ne 0
    $files       = ($InheritanceFlag -band [System.Security.AccessControl.InheritanceFlags]::ObjectInherit) -ne 0
    $inheritOnly = ($PropagationFlag -band [System.Security.AccessControl.PropagationFlags]::InheritOnly) -ne 0
    $noProp      = ($PropagationFlag -band [System.Security.AccessControl.PropagationFlags]::NoPropagateInherit) -ne 0
    if ($inheritOnly) { $thisFolder = $false }
    if ($thisFolder -and $subFolders -and $files -and -not $noProp)      { return "This folder, subfolders and files" }
    elseif ($thisFolder -and $subFolders -and -not $files -and -not $noProp) { return "This folder and subfolders" }
    elseif ($thisFolder -and -not $subFolders -and $files)               { return "This folder and files" }
    elseif ($thisFolder -and -not $subFolders -and -not $files)          { return "This folder only" }
    elseif (-not $thisFolder -and $subFolders -and $files -and -not $noProp)  { return "Subfolders and files" }
    elseif (-not $thisFolder -and $subFolders -and -not $files -and -not $noProp) { return "Subfolders only" }
    elseif (-not $thisFolder -and -not $subFolders -and $files)          { return "Files only" }
    elseif ($thisFolder -and $subFolders -and $files -and $noProp)       { return "This folder, subfolders and files (no propagation)" }
    else {
        $desc = @()
        if ($thisFolder) { $desc += "This folder" }
        if ($subFolders) { $desc += "Subfolders" }
        if ($files)      { $desc += "Files" }
        if ($noProp)     { $desc += "(no propagation)" }
        return ($desc -join ", ")
    }
}

function Get-SimpleRightsDescription {
    param([System.Security.AccessControl.FileSystemRights]$Rights)
    $r  = [int]$Rights
    $FC = 0x1F01FF
    $M  = 0x1301BF
    $RX = 0x1200A9
    $R  = 0x120089
    if (($r -band $FC) -eq $FC) { return "Full Control" }
    if (($r -band $M)  -eq $M)  { return "Modify" }
    $parts = @()
    if (($r -band $RX) -eq $RX)       { $parts += "Read &amp; Execute" }
    elseif (($r -band $R) -eq $R)     { $parts += "Read" }
    if (($r -band 0x000116) -eq 0x000116) { $parts += "Write" }
    if ($parts.Count -gt 0) { return ($parts -join " + ") }
    $specific = @()
    if ($r -band 0x000001) { $specific += "List folder" }
    if ($r -band 0x000002) { $specific += "Add data" }
    if ($r -band 0x000004) { $specific += "Read attributes" }
    if ($r -band 0x000008) { $specific += "Erw. Read attributes" }
    if ($r -band 0x000010) { $specific += "Files erstellen" }
    if ($r -band 0x000020) { $specific += "Create folders" }
    if ($r -band 0x000040) { $specific += "Write attributes" }
    if ($r -band 0x000080) { $specific += "Erw. Write attributes" }
    if ($r -band 0x000100) { $specific += "Subfolders+Files loeschen" }
    if ($r -band 0x010000) { $specific += "Delete" }
    if ($r -band 0x020000) { $specific += "Read permissions" }
    if ($r -band 0x040000) { $specific += "Change permissions" }
    if ($r -band 0x080000) { $specific += "Take ownership" }
    if ($specific.Count -gt 0) { return "Special: " + ($specific -join ", ") }
    return "Special permissions ($r)"
}

function Get-FolderACL {
    param([string]$FolderPath)
    $result = [PSCustomObject]@{
        Path                    = $FolderPath
        Owner                   = ""
        InheritanceEnabled      = $true
        AreAccessRulesProtected = $false
        Rules                   = @()
        Error                   = $null
    }
    try {
        $acl = Get-Acl -Path $FolderPath -ErrorAction Stop
        $result.Owner                   = $acl.Owner
        $result.AreAccessRulesProtected = $acl.AreAccessRulesProtected
        $result.InheritanceEnabled      = -not $acl.AreAccessRulesProtected
        $rules = @()
        foreach ($ace in $acl.Access) {
            $rules += [PSCustomObject]@{
                Identity         = $ace.IdentityReference.Value
                AccessType       = $ace.AccessControlType.ToString()
                Rights           = $ace.FileSystemRights.ToString()
                IsInherited      = $ace.IsInherited
                InheritanceFlags = $ace.InheritanceFlags.ToString()
                PropagationFlags = $ace.PropagationFlags.ToString()
                InheritanceDesc  = Get-InheritanceFlagsDescription -InheritanceFlag $ace.InheritanceFlags -PropagationFlag $ace.PropagationFlags
                RightsSimple     = Get-SimpleRightsDescription -Rights $ace.FileSystemRights
            }
        }
        $result.Rules = $rules | Sort-Object @(
            @{ Expression = { $_.IsInherited };  Ascending = $true },
            @{ Expression = { if ($_.AccessType -eq "Deny") { 0 } else { 1 } }; Ascending = $true },
            @{ Expression = { $_.Identity };     Ascending = $true }
        )
    }
    catch { $result.Error = $_.Exception.Message }
    return $result
}

function Compare-ACLRules {
    param([array]$Rules1, [array]$Rules2)
    if ($Rules1.Count -ne $Rules2.Count) { return $false }
    for ($i = 0; $i -lt $Rules1.Count; $i++) {
        $r1 = $Rules1[$i]; $r2 = $Rules2[$i]
        if ($r1.Identity -ne $r2.Identity -or $r1.AccessType -ne $r2.AccessType -or
            $r1.Rights -ne $r2.Rights -or $r1.IsInherited -ne $r2.IsInherited -or
            $r1.InheritanceFlags -ne $r2.InheritanceFlags -or $r1.PropagationFlags -ne $r2.PropagationFlags) {
            return $false
        }
    }
    return $true
}

function Get-AllFolders {
    param([string]$RootPath, [int]$MaxDepth)
    $allFolders = [System.Collections.Generic.List[string]]::new()
    $allFolders.Add($RootPath)
    $queue = [System.Collections.Queue]::new()
    $queue.Enqueue([PSCustomObject]@{ Path = $RootPath; Depth = 0 })
    while ($queue.Count -gt 0) {
        $current = $queue.Dequeue()
        if ($MaxDepth -gt 0 -and $current.Depth -ge $MaxDepth) { continue }
        try {
            $subDirs = Get-ChildItem -Path $current.Path -Directory -ErrorAction Stop
            foreach ($dir in $subDirs) {
                $allFolders.Add($dir.FullName)
                $queue.Enqueue([PSCustomObject]@{ Path = $dir.FullName; Depth = $current.Depth + 1 })
            }
        } catch { }
    }
    return $allFolders.ToArray()
}



# ============================================================
# MODULE: Report
# ============================================================
# ACLens - Report.psm1
# HTML report generation for NTFS scans and diff reports

function New-HTMLReport {
    param([array]$FolderData, [string]$RootPath, [string]$OutputPath)

    $reportDate   = Get-Date -Format "dd.MM.yyyy HH:mm:ss"
    $computerName = $env:COMPUTERNAME
    $currentUser  = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    $totalFolders   = $FolderData.Count
    $changedFolders = ($FolderData | Where-Object { $_.HasChanges   -eq $true  }).Count
    $errorFolders   = ($FolderData | Where-Object { $null -ne $_.Error }).Count
    $inheritedOnly  = ($FolderData | Where-Object { $_.AllInherited -eq $true -and $_.HasChanges -eq $false }).Count

    $cssB64 = (
        'Cjpyb290IHsKICAtLWJnOiMxRTFFMkU7LS1iZzI6IzJEMkQzRjstLWJnMzojMUExQTJFOy0tYmc0OiMy' +
        'NTI1MzU7CiAgLS1ib3JkZXI6IzRCNTU2MzstLWJvcmRlcjI6IzM3NDE1MTsKICAtLXRleHQ6I0UyRThG' +
        'MDstLXRleHQyOiM5Q0EzQUY7LS10ZXh0MzojNkI3MjgwOwogIC0tYWNjZW50OiNBNzhCRkE7LS1hY2Nl' +
        'bnQyOiM3QzNBRUQ7LS1hY2NlbnQzOiM0QzFEOTU7CiAgLS1ncmVlbjojMzREMzk5Oy0tcmVkOiNGODcx' +
        'NzE7LS15ZWxsb3c6I0ZCQkYyNDstLW9yYW5nZTojRjU5RTBCOwogIC0tZm9udC1tb25vOidDYXNjYWRp' +
        'YSBDb2RlJywnQ29uc29sYXMnLG1vbm9zcGFjZTsKICAtLWZvbnQtdWk6J1NlZ29lIFVJJyxzeXN0ZW0t' +
        'dWksc2Fucy1zZXJpZjsKICAtLXJhZGl1czo2cHg7Cn0KYm9keS5saWdodCB7CiAgLS1iZzojRjlGQUZC' +
        'Oy0tYmcyOiNGM0Y0RjY7LS1iZzM6I0U1RTdFQjstLWJnNDojRDFENURCOwogIC0tYm9yZGVyOiNEMUQ1' +
        'REI7LS1ib3JkZXIyOiNFNUU3RUI7CiAgLS10ZXh0OiMxMTE4Mjc7LS10ZXh0MjojMzc0MTUxOy0tdGV4' +
        'dDM6IzZCNzI4MDsKICAtLWFjY2VudDojN0MzQUVEOy0tYWNjZW50MjojNkQyOEQ5Oy0tYWNjZW50Mzoj' +
        'NEMxRDk1Owp9Cip7Ym94LXNpemluZzpib3JkZXItYm94O21hcmdpbjowO3BhZGRpbmc6MH0KYm9keXti' +
        'YWNrZ3JvdW5kOnZhcigtLWJnKTtjb2xvcjp2YXIoLS10ZXh0KTtmb250LWZhbWlseTp2YXIoLS1mb250' +
        'LXVpKTtmb250LXNpemU6MTNweDtsaW5lLWhlaWdodDoxLjU7dHJhbnNpdGlvbjpiYWNrZ3JvdW5kIC4y' +
        'cyxjb2xvciAuMnN9CgovKiDilIDilIAgSGVhZGVyIOKUgOKUgCAqLwoucmVwb3J0LWhlYWRlcntiYWNr' +
        'Z3JvdW5kOnZhcigtLWJnMyk7Ym9yZGVyLWJvdHRvbToxcHggc29saWQgdmFyKC0tYm9yZGVyKTtwYWRk' +
        'aW5nOjIwcHggMzJweCAxNnB4O2Rpc3BsYXk6ZmxleDthbGlnbi1pdGVtczpjZW50ZXI7anVzdGlmeS1j' +
        'b250ZW50OnNwYWNlLWJldHdlZW47ZmxleC13cmFwOndyYXA7Z2FwOjEycHg7cG9zaXRpb246c3RpY2t5' +
        'O3RvcDowO3otaW5kZXg6MTAwfQouaGVhZGVyLWxlZnR7fQoucmVwb3J0LXRpdGxle2ZvbnQtc2l6ZTox' +
        'OHB4O2ZvbnQtd2VpZ2h0OjcwMDtkaXNwbGF5OmZsZXg7YWxpZ24taXRlbXM6Y2VudGVyO2dhcDo4cHg7' +
        'bWFyZ2luLWJvdHRvbToycHh9Ci5yZXBvcnQtdGl0bGUgLmljb257Y29sb3I6dmFyKC0tYWNjZW50KX0K' +
        'LnJlcG9ydC1zdWJ0aXRsZXtjb2xvcjp2YXIoLS10ZXh0Mik7Zm9udC1mYW1pbHk6dmFyKC0tZm9udC1t' +
        'b25vKTtmb250LXNpemU6MTFweH0KLmhlYWRlci1yaWdodHtkaXNwbGF5OmZsZXg7YWxpZ24taXRlbXM6' +
        'Y2VudGVyO2dhcDoxMHB4O2ZsZXgtc2hyaW5rOjB9Ci50aGVtZS10b2dnbGV7YmFja2dyb3VuZDp2YXIo' +
        'LS1iZzIpO2JvcmRlcjoxcHggc29saWQgdmFyKC0tYm9yZGVyKTtib3JkZXItcmFkaXVzOjIwcHg7cGFk' +
        'ZGluZzo0cHggMTJweDtjdXJzb3I6cG9pbnRlcjtmb250LXNpemU6MTFweDtjb2xvcjp2YXIoLS10ZXh0' +
        'Mik7ZGlzcGxheTpmbGV4O2FsaWduLWl0ZW1zOmNlbnRlcjtnYXA6NnB4O3RyYW5zaXRpb246YWxsIC4x' +
        'NXN9Ci50aGVtZS10b2dnbGU6aG92ZXJ7Ym9yZGVyLWNvbG9yOnZhcigtLWFjY2VudCk7Y29sb3I6dmFy' +
        'KC0tYWNjZW50KX0KCi8qIOKUgOKUgCBTdGF0cyBiYXIg4pSA4pSAICovCi5zdGF0cy1iYXJ7ZGlzcGxh' +
        'eTpmbGV4O2dhcDo4cHg7cGFkZGluZzoxMnB4IDMycHg7YmFja2dyb3VuZDp2YXIoLS1iZzIpO2JvcmRl' +
        'ci1ib3R0b206MXB4IHNvbGlkIHZhcigtLWJvcmRlcik7ZmxleC13cmFwOndyYXB9Ci5zdGF0LWNhcmR7' +
        'YmFja2dyb3VuZDp2YXIoLS1iZzQpO2JvcmRlcjoxcHggc29saWQgdmFyKC0tYm9yZGVyKTtib3JkZXIt' +
        'cmFkaXVzOnZhcigtLXJhZGl1cyk7cGFkZGluZzo4cHggMTZweDtkaXNwbGF5OmZsZXg7ZmxleC1kaXJl' +
        'Y3Rpb246Y29sdW1uO2FsaWduLWl0ZW1zOmNlbnRlcjttaW4td2lkdGg6MTAwcHh9Ci5zdGF0LW51bWJl' +
        'cntmb250LXNpemU6MjBweDtmb250LXdlaWdodDo3MDA7bGluZS1oZWlnaHQ6MX0KLnN0YXQtbGFiZWx7' +
        'Zm9udC1zaXplOjEwcHg7Y29sb3I6dmFyKC0tdGV4dDIpO21hcmdpbi10b3A6MnB4fQouc3RhdC10b3Rh' +
        'bCAuc3RhdC1udW1iZXJ7Y29sb3I6dmFyKC0tYWNjZW50KX0KLnN0YXQtY2hhbmdlZCAuc3RhdC1udW1i' +
        'ZXJ7Y29sb3I6dmFyKC0teWVsbG93KX0KLnN0YXQtaW5oZXJpdGVkIC5zdGF0LW51bWJlcntjb2xvcjp2' +
        'YXIoLS1ncmVlbil9Ci5zdGF0LWVycm9yIC5zdGF0LW51bWJlcntjb2xvcjp2YXIoLS1yZWQpfQoKLyog' +
        '4pSA4pSAIFRvb2xiYXIg4pSA4pSAICovCi50b29sYmFye2Rpc3BsYXk6ZmxleDthbGlnbi1pdGVtczpj' +
        'ZW50ZXI7Z2FwOjhweDtwYWRkaW5nOjEwcHggMzJweDtiYWNrZ3JvdW5kOnZhcigtLWJnMyk7Ym9yZGVy' +
        'LWJvdHRvbToxcHggc29saWQgdmFyKC0tYm9yZGVyMik7ZmxleC13cmFwOndyYXA7cG9zaXRpb246c3Rp' +
        'Y2t5O3RvcDo2MXB4O3otaW5kZXg6OTl9Ci5maWx0ZXItZ3JvdXB7ZGlzcGxheTpmbGV4O2dhcDo0cHh9' +
        'Ci5idG57YmFja2dyb3VuZDp2YXIoLS1iZzIpO2JvcmRlcjoxcHggc29saWQgdmFyKC0tYm9yZGVyKTti' +
        'b3JkZXItcmFkaXVzOnZhcigtLXJhZGl1cyk7Y29sb3I6dmFyKC0tdGV4dCk7Y3Vyc29yOnBvaW50ZXI7' +
        'Zm9udC1zaXplOjExcHg7cGFkZGluZzo0cHggMTJweDtmb250LWZhbWlseTp2YXIoLS1mb250LXVpKTt0' +
        'cmFuc2l0aW9uOmFsbCAuMTVzfQouYnRuOmhvdmVye2JhY2tncm91bmQ6dmFyKC0tYmc0KTtib3JkZXIt' +
        'Y29sb3I6dmFyKC0tYWNjZW50KX0KLmJ0bi1hY2NlbnR7YmFja2dyb3VuZDp2YXIoLS1hY2NlbnQyKTti' +
        'b3JkZXItY29sb3I6dmFyKC0tYWNjZW50Myk7Y29sb3I6I2ZmZn0KLmJ0bi1hY2NlbnQ6aG92ZXJ7YmFj' +
        'a2dyb3VuZDp2YXIoLS1hY2NlbnQzKX0KLmZpbHRlci1idG57cGFkZGluZzozcHggMTBweDtib3JkZXIt' +
        'cmFkaXVzOjIwcHh9Ci5maWx0ZXItYnRuLmFjdGl2ZXtiYWNrZ3JvdW5kOnZhcigtLWFjY2VudDIpO2Jv' +
        'cmRlci1jb2xvcjp2YXIoLS1hY2NlbnQzKTtjb2xvcjojZmZmfQouc2VhcmNoLWJveHtiYWNrZ3JvdW5k' +
        'OnZhcigtLWJnMik7Ym9yZGVyOjFweCBzb2xpZCB2YXIoLS1ib3JkZXIpO2JvcmRlci1yYWRpdXM6dmFy' +
        'KC0tcmFkaXVzKTtjb2xvcjp2YXIoLS10ZXh0KTtmb250LXNpemU6MTFweDtwYWRkaW5nOjRweCAxMHB4' +
        'O3dpZHRoOjIyMHB4O291dGxpbmU6bm9uZTtmb250LWZhbWlseTp2YXIoLS1mb250LXVpKX0KLnNlYXJj' +
        'aC1ib3g6Zm9jdXN7Ym9yZGVyLWNvbG9yOnZhcigtLWFjY2VudCl9Ci5zZWFyY2gtYm94OjpwbGFjZWhv' +
        'bGRlcntjb2xvcjp2YXIoLS10ZXh0Myl9Ci50b29sYmFyLXNwYWNlcntmbGV4OjF9CgovKiDilIDilIAg' +
        'UGFnZSBuYXZpZ2F0aW9uIOKUgOKUgCAqLwoucGFnZS1uYXZ7ZGlzcGxheTpmbGV4O2FsaWduLWl0ZW1z' +
        'OmNlbnRlcjtnYXA6NnB4O3BhZGRpbmc6MTBweCAzMnB4O2JhY2tncm91bmQ6dmFyKC0tYmcpO2JvcmRl' +
        'ci1ib3R0b206MXB4IHNvbGlkIHZhcigtLWJvcmRlcjIpO2ZsZXgtd3JhcDp3cmFwfQoucGFnZS1uYXYt' +
        'bGFiZWx7Zm9udC1zaXplOjExcHg7Y29sb3I6dmFyKC0tdGV4dDIpfQoucGFnZS1idG57YmFja2dyb3Vu' +
        'ZDp2YXIoLS1iZzIpO2JvcmRlcjoxcHggc29saWQgdmFyKC0tYm9yZGVyKTtib3JkZXItcmFkaXVzOjRw' +
        'eDtjb2xvcjp2YXIoLS10ZXh0KTtjdXJzb3I6cG9pbnRlcjtmb250LXNpemU6MTFweDtwYWRkaW5nOjNw' +
        'eCAxMHB4O2ZvbnQtZmFtaWx5OnZhcigtLWZvbnQtdWkpO21pbi13aWR0aDozMnB4O3RleHQtYWxpZ246' +
        'Y2VudGVyfQoucGFnZS1idG46aG92ZXJ7Ym9yZGVyLWNvbG9yOnZhcigtLWFjY2VudCk7Y29sb3I6dmFy' +
        'KC0tYWNjZW50KX0KLnBhZ2UtYnRuLmFjdGl2ZXtiYWNrZ3JvdW5kOnZhcigtLWFjY2VudDIpO2JvcmRl' +
        'ci1jb2xvcjp2YXIoLS1hY2NlbnQzKTtjb2xvcjojZmZmfQoucGFnZS1idG46ZGlzYWJsZWR7b3BhY2l0' +
        'eTouMztjdXJzb3I6ZGVmYXVsdH0KLnBhZ2Utc2l6ZS1zZWxlY3R7YmFja2dyb3VuZDp2YXIoLS1iZzIp' +
        'O2JvcmRlcjoxcHggc29saWQgdmFyKC0tYm9yZGVyKTtib3JkZXItcmFkaXVzOjRweDtjb2xvcjp2YXIo' +
        'LS10ZXh0KTtmb250LXNpemU6MTFweDtwYWRkaW5nOjNweCA2cHg7Zm9udC1mYW1pbHk6dmFyKC0tZm9u' +
        'dC11aSl9Ci5wYWdlLWluZm97Zm9udC1zaXplOjExcHg7Y29sb3I6dmFyKC0tdGV4dDIpO21hcmdpbi1s' +
        'ZWZ0OmF1dG99CgovKiDilIDilIAgRm9sZGVyIGxpc3Qg4pSA4pSAICovCi5mb2xkZXItbGlzdHtwYWRk' +
        'aW5nOjEwcHggMjRweH0KCi8qIOKUgOKUgCBGb2xkZXIgcm93IOKUgOKUgCAqLwouZm9sZGVyLXJvd3ti' +
        'b3JkZXI6MXB4IHNvbGlkIHZhcigtLWJvcmRlcjIpO2JvcmRlci1yYWRpdXM6dmFyKC0tcmFkaXVzKTtt' +
        'YXJnaW4tYm90dG9tOjNweDtvdmVyZmxvdzpoaWRkZW47dHJhbnNpdGlvbjpib3JkZXItY29sb3IgLjE1' +
        'c30KLmZvbGRlci1yb3c6aG92ZXJ7Ym9yZGVyLWNvbG9yOnZhcigtLWJvcmRlcil9Ci5mb2xkZXItcm93' +
        'LnN0YXR1cy1yb290e2JvcmRlci1sZWZ0OjNweCBzb2xpZCB2YXIoLS1hY2NlbnQpO2JhY2tncm91bmQ6' +
        'cmdiYSgxNjcsMTM5LDI1MCwuMDQpfQouZm9sZGVyLXJvdy5zdGF0dXMtY2hhbmdlZHtib3JkZXItbGVm' +
        'dDozcHggc29saWQgdmFyKC0teWVsbG93KTtiYWNrZ3JvdW5kOnJnYmEoMjUxLDE5MSwzNiwuMDQpfQou' +
        'Zm9sZGVyLXJvdy5zdGF0dXMtZXJyb3J7Ym9yZGVyLWxlZnQ6M3B4IHNvbGlkIHZhcigtLXJlZCl9Ci5m' +
        'b2xkZXItcm93LnN0YXR1cy1pbmhlcml0ZWR7Ym9yZGVyLWxlZnQ6M3B4IHNvbGlkIHZhcigtLWJvcmRl' +
        'cjIpfQoKLmZvbGRlci1oZWFkZXJ7ZGlzcGxheTpmbGV4O2FsaWduLWl0ZW1zOmNlbnRlcjtqdXN0aWZ5' +
        'LWNvbnRlbnQ6c3BhY2UtYmV0d2VlbjtwYWRkaW5nOjdweCAxNHB4O2N1cnNvcjpwb2ludGVyO3VzZXIt' +
        'c2VsZWN0Om5vbmU7Z2FwOjhweH0KLmZvbGRlci1oZWFkZXI6aG92ZXJ7YmFja2dyb3VuZDpyZ2JhKDE2' +
        'NywxMzksMjUwLC4wNSl9Ci5mb2xkZXItbGVmdHtkaXNwbGF5OmZsZXg7YWxpZ24taXRlbXM6Y2VudGVy' +
        'O2dhcDo3cHg7ZmxleDoxO21pbi13aWR0aDowfQouZm9sZGVyLXJpZ2h0e2Rpc3BsYXk6ZmxleDthbGln' +
        'bi1pdGVtczpjZW50ZXI7Z2FwOjZweDtmbGV4LXNocmluazowfQoudG9nZ2xlLWljb257Y29sb3I6dmFy' +
        'KC0tdGV4dDMpO2ZvbnQtc2l6ZToxMHB4O3RyYW5zaXRpb246dHJhbnNmb3JtIC4ycztmbGV4LXNocmlu' +
        'azowfQouZm9sZGVyLWljb257Zm9udC1zaXplOjE0cHg7ZmxleC1zaHJpbms6MH0KLmZvbGRlci1wYXRo' +
        'e2ZvbnQtZmFtaWx5OnZhcigtLWZvbnQtbW9ubyk7Zm9udC1zaXplOjEycHg7d2hpdGUtc3BhY2U6bm93' +
        'cmFwO292ZXJmbG93OmhpZGRlbjt0ZXh0LW92ZXJmbG93OmVsbGlwc2lzfQoKLnN0YXR1cy1iYWRnZSwu' +
        'aW5oZXJpdC1iYWRnZSwub3duZXItYmFkZ2V7Zm9udC1zaXplOjEwcHg7cGFkZGluZzoycHggN3B4O2Jv' +
        'cmRlci1yYWRpdXM6MTBweDt3aGl0ZS1zcGFjZTpub3dyYXA7Zm9udC13ZWlnaHQ6NTAwfQoub3duZXIt' +
        'YmFkZ2V7YmFja2dyb3VuZDp2YXIoLS1iZzQpO2JvcmRlcjoxcHggc29saWQgdmFyKC0tYm9yZGVyMik7' +
        'Y29sb3I6dmFyKC0tdGV4dDIpfQouc3RhdHVzLWJhZGdlLnN0YXR1cy1yb290e2JhY2tncm91bmQ6cmdi' +
        'YSgxNjcsMTM5LDI1MCwuMTUpO2JvcmRlcjoxcHggc29saWQgdmFyKC0tYWNjZW50Mik7Y29sb3I6dmFy' +
        'KC0tYWNjZW50KX0KLnN0YXR1cy1iYWRnZS5zdGF0dXMtY2hhbmdlZHtiYWNrZ3JvdW5kOnJnYmEoMjUx' +
        'LDE5MSwzNiwuMTUpO2JvcmRlcjoxcHggc29saWQgdmFyKC0teWVsbG93KTtjb2xvcjp2YXIoLS15ZWxs' +
        'b3cpfQouc3RhdHVzLWJhZGdlLnN0YXR1cy1lcnJvcntiYWNrZ3JvdW5kOnJnYmEoMjQxLDExMywxMTMs' +
        'LjEyKTtib3JkZXI6MXB4IHNvbGlkIHZhcigtLXJlZCk7Y29sb3I6dmFyKC0tcmVkKX0KLnN0YXR1cy1i' +
        'YWRnZS5zdGF0dXMtaW5oZXJpdGVke2JhY2tncm91bmQ6cmdiYSg1MiwyMTEsMTUzLC4xMik7Ym9yZGVy' +
        'OjFweCBzb2xpZCByZ2JhKDUyLDIxMSwxNTMsLjMpO2NvbG9yOnZhcigtLWdyZWVuKX0KLmluaGVyaXQt' +
        'eWVze2JhY2tncm91bmQ6cmdiYSg1MiwyMTEsMTUzLC4xMik7Ym9yZGVyOjFweCBzb2xpZCByZ2JhKDUy' +
        'LDIxMSwxNTMsLjMpO2NvbG9yOnZhcigtLWdyZWVuKX0KLmluaGVyaXQtbm97YmFja2dyb3VuZDpyZ2Jh' +
        'KDI0MSwxMTMsMTEzLC4xMik7Ym9yZGVyOjFweCBzb2xpZCByZ2JhKDI0MSwxMTMsMTEzLC4zKTtjb2xv' +
        'cjp2YXIoLS1yZWQpfQoKLmZvbGRlci1kZXRhaWxze2Rpc3BsYXk6bm9uZTtwYWRkaW5nOjEwcHggMTZw' +
        'eCAxNHB4O2JvcmRlci10b3A6MXB4IHNvbGlkIHZhcigtLWJvcmRlcjIpO2JhY2tncm91bmQ6dmFyKC0t' +
        'YmcpfQouZm9sZGVyLWRldGFpbHMub3BlbntkaXNwbGF5OmJsb2NrfQouZGV0YWlscy1tZXRhe2Rpc3Bs' +
        'YXk6ZmxleDtmbGV4LXdyYXA6d3JhcDtnYXA6MTRweDttYXJnaW4tYm90dG9tOjEwcHg7cGFkZGluZzo4' +
        'cHggMTJweDtiYWNrZ3JvdW5kOnZhcigtLWJnMik7Ym9yZGVyLXJhZGl1czp2YXIoLS1yYWRpdXMpO2Jv' +
        'cmRlcjoxcHggc29saWQgdmFyKC0tYm9yZGVyMik7Zm9udC1zaXplOjExcHg7Y29sb3I6dmFyKC0tdGV4' +
        'dDIpfQouZGV0YWlscy1tZXRhIHN0cm9uZ3tjb2xvcjp2YXIoLS10ZXh0KX0KLnNhbWUtYXMtcGFyZW50' +
        'e2NvbG9yOnZhcigtLXRleHQzKTtmb250LXN0eWxlOml0YWxpYztmb250LXNpemU6MTFweDtwYWRkaW5n' +
        'OjZweCAxMHB4O2JhY2tncm91bmQ6dmFyKC0tYmcyKTtib3JkZXItcmFkaXVzOnZhcigtLXJhZGl1cyk7' +
        'Ym9yZGVyOjFweCBkYXNoZWQgdmFyKC0tYm9yZGVyMil9Ci5lcnJvci1ib3h7YmFja2dyb3VuZDpyZ2Jh' +
        'KDI0MSwxMTMsMTEzLC4wNyk7Ym9yZGVyOjFweCBzb2xpZCB2YXIoLS1yZWQpO2JvcmRlci1yYWRpdXM6' +
        'dmFyKC0tcmFkaXVzKTtjb2xvcjp2YXIoLS1yZWQpO3BhZGRpbmc6OHB4IDEycHg7Zm9udC1zaXplOjEx' +
        'cHh9CgoucnVsZXMtc2VjdGlvbnttYXJnaW4tYm90dG9tOjhweH0KLnJ1bGVzLXNlY3Rpb24tdGl0bGV7' +
        'Zm9udC1zaXplOjEwcHg7Zm9udC13ZWlnaHQ6NzAwO3RleHQtdHJhbnNmb3JtOnVwcGVyY2FzZTtsZXR0' +
        'ZXItc3BhY2luZzouNnB4O3BhZGRpbmc6MnB4IDhweDttYXJnaW4tYm90dG9tOjRweDtib3JkZXItcmFk' +
        'aXVzOjNweDtkaXNwbGF5OmlubGluZS1ibG9ja30KLmV4cGxpY2l0LXRpdGxle2NvbG9yOnZhcigtLWFj' +
        'Y2VudCk7YmFja2dyb3VuZDpyZ2JhKDE2NywxMzksMjUwLC4xMil9Ci5pbmhlcml0ZWQtdGl0bGV7Y29s' +
        'b3I6dmFyKC0tdGV4dDMpO2JhY2tncm91bmQ6dmFyKC0tYmcyKX0KLnJ1bGVzLXRhYmxle3dpZHRoOjEw' +
        'MCU7Ym9yZGVyLWNvbGxhcHNlOmNvbGxhcHNlO2ZvbnQtc2l6ZToxMXB4fQoucnVsZXMtdGFibGUgdGhl' +
        'YWQgdHJ7YmFja2dyb3VuZDp2YXIoLS1iZzQpfQoucnVsZXMtdGFibGUgdGh7dGV4dC1hbGlnbjpsZWZ0' +
        'O3BhZGRpbmc6NHB4IDhweDtjb2xvcjp2YXIoLS10ZXh0Mik7Zm9udC13ZWlnaHQ6NTAwO2ZvbnQtc2l6' +
        'ZToxMHB4O2JvcmRlci1ib3R0b206MXB4IHNvbGlkIHZhcigtLWJvcmRlcil9Ci5ydWxlcy10YWJsZSB0' +
        'ZHtwYWRkaW5nOjRweCA4cHg7Ym9yZGVyLWJvdHRvbToxcHggc29saWQgdmFyKC0tYm9yZGVyMik7dmVy' +
        'dGljYWwtYWxpZ246bWlkZGxlfQoucnVsZXMtdGFibGUgdGJvZHkgdHI6aG92ZXJ7YmFja2dyb3VuZDpy' +
        'Z2JhKDE2NywxMzksMjUwLC4wNCl9Ci5ydWxlcy10YWJsZSB0Ym9keSB0cjpsYXN0LWNoaWxkIHRke2Jv' +
        'cmRlci1ib3R0b206bm9uZX0KLmluaGVyaXRlZC10YWJsZXtvcGFjaXR5Oi43fQouZGVueS1ydWxle2Jh' +
        'Y2tncm91bmQ6cmdiYSgyNDEsMTEzLDExMywuMDYpIWltcG9ydGFudH0KLmRlbnktcnVsZSB0ZHtjb2xv' +
        'cjp2YXIoLS1yZWQpIWltcG9ydGFudH0KLmlkZW50aXR5e2ZvbnQtZmFtaWx5OnZhcigtLWZvbnQtbW9u' +
        'byk7Zm9udC1zaXplOjEwcHg7Zm9udC13ZWlnaHQ6NTAwfQouc2NvcGV7Y29sb3I6dmFyKC0tdGV4dDIp' +
        'O2ZvbnQtc2l6ZToxMHB4fQoubm8tcnVsZXN7Y29sb3I6dmFyKC0tdGV4dDMpO2ZvbnQtc3R5bGU6aXRh' +
        'bGljO2ZvbnQtc2l6ZToxMXB4O3BhZGRpbmc6NHB4IDhweH0KLmJhZGdle2Rpc3BsYXk6aW5saW5lLWJs' +
        'b2NrO2ZvbnQtc2l6ZTo5cHg7cGFkZGluZzoxcHggNnB4O2JvcmRlci1yYWRpdXM6OHB4O2ZvbnQtd2Vp' +
        'Z2h0OjUwMDt3aGl0ZS1zcGFjZTpub3dyYXB9Ci5iYWRnZS1hbGxvd3tiYWNrZ3JvdW5kOnJnYmEoNTIs' +
        'MjExLDE1MywuMTIpO2JvcmRlcjoxcHggc29saWQgcmdiYSg1MiwyMTEsMTUzLC4zNSk7Y29sb3I6dmFy' +
        'KC0tZ3JlZW4pfQouYmFkZ2UtZGVueXtiYWNrZ3JvdW5kOnJnYmEoMjQxLDExMywxMTMsLjEyKTtib3Jk' +
        'ZXI6MXB4IHNvbGlkIHJnYmEoMjQxLDExMywxMTMsLjM1KTtjb2xvcjp2YXIoLS1yZWQpfQouYmFkZ2Ut' +
        'aW5oZXJpdGVke2JhY2tncm91bmQ6dmFyKC0tYmcyKTtib3JkZXI6MXB4IHNvbGlkIHZhcigtLWJvcmRl' +
        'cjIpO2NvbG9yOnZhcigtLXRleHQzKX0KLmJhZGdlLWV4cGxpY2l0e2JhY2tncm91bmQ6cmdiYSgxNjcs' +
        'MTM5LDI1MCwuMTUpO2JvcmRlcjoxcHggc29saWQgcmdiYSgxNjcsMTM5LDI1MCwuMyk7Y29sb3I6dmFy' +
        'KC0tYWNjZW50KX0KCi8qIOKUgOKUgCBBZ2VuZGEg4pSA4pSAICovCi5hZ2VuZGEtYmFye2JhY2tncm91' +
        'bmQ6dmFyKC0tYmczKTtib3JkZXItYm90dG9tOjFweCBzb2xpZCB2YXIoLS1ib3JkZXIpfQouYWdlbmRh' +
        'LXRvZ2dsZXt3aWR0aDoxMDAlO2Rpc3BsYXk6ZmxleDthbGlnbi1pdGVtczpjZW50ZXI7Z2FwOjhweDtw' +
        'YWRkaW5nOjhweCAzMnB4O2JhY2tncm91bmQ6bm9uZTtib3JkZXI6bm9uZTtjb2xvcjp2YXIoLS10ZXh0' +
        'Mik7Zm9udC1zaXplOjExcHg7Zm9udC1mYW1pbHk6dmFyKC0tZm9udC11aSk7Y3Vyc29yOnBvaW50ZXI7' +
        'dGV4dC1hbGlnbjpsZWZ0fQouYWdlbmRhLXRvZ2dsZTpob3ZlcntiYWNrZ3JvdW5kOnJnYmEoMTY3LDEz' +
        'OSwyNTAsLjA2KTtjb2xvcjp2YXIoLS1hY2NlbnQpfQouYWdlbmRhLXRvZ2dsZS1pY29ue2ZvbnQtc2l6' +
        'ZTo5cHg7dHJhbnNpdGlvbjp0cmFuc2Zvcm0gLjJzfQouYWdlbmRhLXRvZ2dsZS1oaW50e21hcmdpbi1s' +
        'ZWZ0OmF1dG87Zm9udC1zaXplOjEwcHg7Y29sb3I6dmFyKC0tdGV4dDMpfQouYWdlbmRhLWNvbnRlbnR7' +
        'cGFkZGluZzoxNHB4IDMycHggMThweDtib3JkZXItdG9wOjFweCBzb2xpZCB2YXIoLS1ib3JkZXIyKTti' +
        'YWNrZ3JvdW5kOnZhcigtLWJnKX0KLmFnZW5kYS1ncmlke2Rpc3BsYXk6Z3JpZDtncmlkLXRlbXBsYXRl' +
        'LWNvbHVtbnM6cmVwZWF0KGF1dG8tZmlsbCxtaW5tYXgoMjgwcHgsMWZyKSk7Z2FwOjEwcHh9Ci5hZ2Vu' +
        'ZGEtc2VjdGlvbntiYWNrZ3JvdW5kOnZhcigtLWJnNCk7Ym9yZGVyOjFweCBzb2xpZCB2YXIoLS1ib3Jk' +
        'ZXIyKTtib3JkZXItcmFkaXVzOnZhcigtLXJhZGl1cyk7cGFkZGluZzoxMHB4IDEycHh9Ci5hZ2VuZGEt' +
        'c2VjdGlvbi10aXRsZXtmb250LXNpemU6OXB4O2ZvbnQtd2VpZ2h0OjcwMDt0ZXh0LXRyYW5zZm9ybTp1' +
        'cHBlcmNhc2U7bGV0dGVyLXNwYWNpbmc6LjhweDtjb2xvcjp2YXIoLS1hY2NlbnQpO21hcmdpbi1ib3R0' +
        'b206NnB4O3BhZGRpbmctYm90dG9tOjRweDtib3JkZXItYm90dG9tOjFweCBzb2xpZCB2YXIoLS1ib3Jk' +
        'ZXIyKX0KLmFnZW5kYS1yb3d7ZGlzcGxheTpmbGV4O2FsaWduLWl0ZW1zOmJhc2VsaW5lO2dhcDo2cHg7' +
        'bWFyZ2luLWJvdHRvbTo0cHg7Zm9udC1zaXplOjEwcHg7bGluZS1oZWlnaHQ6MS40fQouYWdlbmRhLXJv' +
        'dzpsYXN0LWNoaWxke21hcmdpbi1ib3R0b206MH0KLmFnZW5kYS1kb3R7d2lkdGg6N3B4O2hlaWdodDo3' +
        'cHg7Ym9yZGVyLXJhZGl1czoycHg7ZmxleC1zaHJpbms6MDttYXJnaW4tdG9wOjNweDtkaXNwbGF5Omlu' +
        'bGluZS1ibG9ja30KLmFnZW5kYS1sYWJlbHtjb2xvcjp2YXIoLS10ZXh0KTtmb250LXdlaWdodDo2MDA7' +
        'd2hpdGUtc3BhY2U6bm93cmFwO2ZsZXgtc2hyaW5rOjB9Ci5hZ2VuZGEtcGVybXtmb250LWZhbWlseTp2' +
        'YXIoLS1mb250LW1vbm8pO2ZvbnQtc2l6ZTo5cHg7Y29sb3I6dmFyKC0tdGV4dDIpO2ZvbnQtd2VpZ2h0' +
        'OjUwMDt3aGl0ZS1zcGFjZTpub3dyYXA7ZmxleC1zaHJpbms6MH0KLmFnZW5kYS1kZXNje2NvbG9yOnZh' +
        'cigtLXRleHQyKTtmb250LXNpemU6MTBweH0KLmFnZW5kYS1iYWRnZXtmbGV4LXNocmluazowfQoKLyog' +
        '4pSA4pSAIEZvb3RlciDilIDilIAgKi8KLnJlcG9ydC1mb290ZXJ7dGV4dC1hbGlnbjpjZW50ZXI7cGFk' +
        'ZGluZzoxNHB4O2NvbG9yOnZhcigtLXRleHQzKTtmb250LXNpemU6MTBweDtib3JkZXItdG9wOjFweCBz' +
        'b2xpZCB2YXIoLS1ib3JkZXIyKTtiYWNrZ3JvdW5kOnZhcigtLWJnMyk7bWFyZ2luLXRvcDoxNnB4fQoK' +
        'Lyog4pSA4pSAIEVtcHR5IHN0YXRlIOKUgOKUgCAqLwouZW1wdHktcGFnZXtwYWRkaW5nOjYwcHggMzJw' +
        'eDt0ZXh0LWFsaWduOmNlbnRlcjtjb2xvcjp2YXIoLS10ZXh0Myl9Ci5lbXB0eS1wYWdlIC5pY29ue2Zv' +
        'bnQtc2l6ZTo0OHB4O21hcmdpbi1ib3R0b206MTJweH0KLmVtcHR5LXBhZ2UgcHtmb250LXNpemU6MTRw' +
        'eH0KCkBtZWRpYSBwcmludHsudG9vbGJhciwucGFnZS1uYXYsLnN0YXRzLWJhciwuYWdlbmRhLWJhciwu' +
        'dGhlbWUtdG9nZ2xle2Rpc3BsYXk6bm9uZX0uZm9sZGVyLWRldGFpbHN7ZGlzcGxheTpibG9jayFpbXBv' +
        'cnRhbnR9LmZvbGRlci1yb3d7YnJlYWstaW5zaWRlOmF2b2lkfX0K'
    )
    $jsB64 = (
        'Ci8vIFRoZW1lIHRvZ2dsZQpmdW5jdGlvbiB0b2dnbGVUaGVtZSgpewogIGRvY3VtZW50LmJvZHkuY2xh' +
        'c3NMaXN0LnRvZ2dsZSgnbGlnaHQnKTsKICB2YXIgYnRuPWRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCd0' +
        'aGVtZUJ0bicpOwogIGJ0bi50ZXh0Q29udGVudD1kb2N1bWVudC5ib2R5LmNsYXNzTGlzdC5jb250YWlu' +
        'cygnbGlnaHQnKT8n8J+MmSBEYXJrJzon4piA77iPIExpZ2h0JzsKICBsb2NhbFN0b3JhZ2Uuc2V0SXRl' +
        'bSgnYWNsLXRoZW1lJyxkb2N1bWVudC5ib2R5LmNsYXNzTGlzdC5jb250YWlucygnbGlnaHQnKT8nbGln' +
        'aHQnOidkYXJrJyk7Cn0KKGZ1bmN0aW9uKCl7CiAgdmFyIHQ9bG9jYWxTdG9yYWdlLmdldEl0ZW0oJ2Fj' +
        'bC10aGVtZScpOwogIGlmKHQ9PT0nbGlnaHQnKXtkb2N1bWVudC5ib2R5LmNsYXNzTGlzdC5hZGQoJ2xp' +
        'Z2h0Jyk7dmFyIGI9ZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3RoZW1lQnRuJyk7aWYoYiliLnRleHRD' +
        'b250ZW50PSfwn4yZIERhcmsnO30KfSkoKTsKCi8vIEFnZW5kYQpmdW5jdGlvbiB0b2dnbGVBZ2VuZGEo' +
        'YnRuKXsKICB2YXIgYz1kb2N1bWVudC5nZXRFbGVtZW50QnlJZCgnYWdlbmRhQ29udGVudCcpLGljPWJ0' +
        'bi5xdWVyeVNlbGVjdG9yKCcuYWdlbmRhLXRvZ2dsZS1pY29uJyksaGludD1idG4ucXVlcnlTZWxlY3Rv' +
        'cignLmFnZW5kYS10b2dnbGUtaGludCcpLGhpZGRlbj1jLnN0eWxlLmRpc3BsYXk9PT0nJ3x8Yy5zdHls' +
        'ZS5kaXNwbGF5PT09J25vbmUnOwogIGMuc3R5bGUuZGlzcGxheT1oaWRkZW4/J2Jsb2NrJzonbm9uZSc7' +
        'CiAgaWMuc3R5bGUudHJhbnNmb3JtPWhpZGRlbj8ncm90YXRlKDkwZGVnKSc6Jyc7CiAgaGludC50ZXh0' +
        'Q29udGVudD1oaWRkZW4/J2NsaWNrIHRvIGNvbGxhcHNlJzonY2xpY2sgdG8gZXhwYW5kJzsKfQoKLy8g' +
        'Rm9sZGVyIHRvZ2dsZQpmdW5jdGlvbiB0b2dnbGUoaWQsaGRyKXsKICB2YXIgZD1kb2N1bWVudC5nZXRF' +
        'bGVtZW50QnlJZChpZCksaWM9aGRyLnF1ZXJ5U2VsZWN0b3IoJy50b2dnbGUtaWNvbicpOwogIGlmKGQu' +
        'Y2xhc3NMaXN0LmNvbnRhaW5zKCdvcGVuJykpe2QuY2xhc3NMaXN0LnJlbW92ZSgnb3BlbicpO2ljLnN0' +
        'eWxlLnRyYW5zZm9ybT0nJzt9CiAgZWxzZXtkLmNsYXNzTGlzdC5hZGQoJ29wZW4nKTtpYy5zdHlsZS50' +
        'cmFuc2Zvcm09J3JvdGF0ZSg5MGRlZyknO30KfQoKLy8g4pSA4pSAIFBhZ2luYXRpb24gZW5naW5lIOKU' +
        'gOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKU' +
        'gOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgOKUgAp2YXIgX2Fs' +
        'bFJvd3M9W107CnZhciBfZmlsdGVyZWQ9W107CnZhciBfcGFnZT0xOwp2YXIgX3BhZ2VTaXplPTI1Owp2' +
        'YXIgX2ZpbHRlcj0nYWxsJzsKdmFyIF9zZWFyY2g9Jyc7CgpmdW5jdGlvbiBpbml0KCl7CiAgX2FsbFJv' +
        'd3M9QXJyYXkuZnJvbShkb2N1bWVudC5xdWVyeVNlbGVjdG9yQWxsKCcuZm9sZGVyLXJvdycpKTsKICBy' +
        'ZW5kZXJQYWdlQnV0dG9ucygpOwogIHNob3dQYWdlKDEpOwp9CgpmdW5jdGlvbiBmaWx0ZXJSb3dzKGYs' +
        'YnRuKXsKICBfZmlsdGVyPWY7CiAgZG9jdW1lbnQucXVlcnlTZWxlY3RvckFsbCgnLmZpbHRlci1idG4n' +
        'KS5mb3JFYWNoKGZ1bmN0aW9uKGIpe2IuY2xhc3NMaXN0LnJlbW92ZSgnYWN0aXZlJyk7fSk7CiAgYnRu' +
        'LmNsYXNzTGlzdC5hZGQoJ2FjdGl2ZScpOwogIF9wYWdlPTE7CiAgYXBwbHlGaWx0ZXIoKTsKfQoKZnVu' +
        'Y3Rpb24gc2VhcmNoUm93cyh2KXsKICBfc2VhcmNoPXYudG9Mb3dlckNhc2UoKTsKICBfcGFnZT0xOwog' +
        'IGFwcGx5RmlsdGVyKCk7Cn0KCmZ1bmN0aW9uIGFwcGx5RmlsdGVyKCl7CiAgX2ZpbHRlcmVkPV9hbGxS' +
        'b3dzLmZpbHRlcihmdW5jdGlvbihyb3cpewogICAgdmFyIG1hdGNoRmlsdGVyPXRydWU7CiAgICBpZihf' +
        'ZmlsdGVyPT09J2NoYW5nZWQnKSAgIG1hdGNoRmlsdGVyPXJvdy5jbGFzc0xpc3QuY29udGFpbnMoJ3N0' +
        'YXR1cy1jaGFuZ2VkJyk7CiAgICBpZihfZmlsdGVyPT09J2luaGVyaXRlZCcpIG1hdGNoRmlsdGVyPXJv' +
        'dy5jbGFzc0xpc3QuY29udGFpbnMoJ3N0YXR1cy1pbmhlcml0ZWQnKTsKICAgIGlmKF9maWx0ZXI9PT0n' +
        'ZXJyb3InKSAgICAgbWF0Y2hGaWx0ZXI9cm93LmNsYXNzTGlzdC5jb250YWlucygnc3RhdHVzLWVycm9y' +
        'Jyk7CiAgICB2YXIgbWF0Y2hTZWFyY2g9IV9zZWFyY2h8fHJvdy50ZXh0Q29udGVudC50b0xvd2VyQ2Fz' +
        'ZSgpLmluZGV4T2YoX3NlYXJjaCk+PTA7CiAgICByZXR1cm4gbWF0Y2hGaWx0ZXImJm1hdGNoU2VhcmNo' +
        'OwogIH0pOwogIHJlbmRlclBhZ2VCdXR0b25zKCk7CiAgc2hvd1BhZ2UoMSk7Cn0KCmZ1bmN0aW9uIHNl' +
        'dFBhZ2VTaXplKHZhbCl7CiAgX3BhZ2VTaXplPXBhcnNlSW50KHZhbCk7CiAgX3BhZ2U9MTsKICByZW5k' +
        'ZXJQYWdlQnV0dG9ucygpOwogIHNob3dQYWdlKDEpOwp9CgpmdW5jdGlvbiBzaG93UGFnZShwKXsKICBf' +
        'cGFnZT1wOwogIHZhciB0b3RhbD1NYXRoLmNlaWwoX2ZpbHRlcmVkLmxlbmd0aC9fcGFnZVNpemUpfHwx' +
        'OwogIGlmKF9wYWdlPDEpX3BhZ2U9MTsKICBpZihfcGFnZT50b3RhbClfcGFnZT10b3RhbDsKCiAgLy8g' +
        'SGlkZSBhbGwgcm93cyBmaXJzdAogIF9hbGxSb3dzLmZvckVhY2goZnVuY3Rpb24ocil7ci5zdHlsZS5k' +
        'aXNwbGF5PSdub25lJzt9KTsKCiAgLy8gU2hvdyBjdXJyZW50IHBhZ2Ugc2xpY2UKICB2YXIgc3RhcnQ9' +
        'KF9wYWdlLTEpKl9wYWdlU2l6ZTsKICB2YXIgZW5kPU1hdGgubWluKHN0YXJ0K19wYWdlU2l6ZSxfZmls' +
        'dGVyZWQubGVuZ3RoKTsKICBmb3IodmFyIGk9c3RhcnQ7aTxlbmQ7aSsrKXtfZmlsdGVyZWRbaV0uc3R5' +
        'bGUuZGlzcGxheT0nJzt9CgogIC8vIFVwZGF0ZSBwYWdlIGluZm8KICB2YXIgaW5mbz1kb2N1bWVudC5n' +
        'ZXRFbGVtZW50QnlJZCgncGFnZUluZm8nKTsKICBpZihpbmZvKXsKICAgIGlmKF9maWx0ZXJlZC5sZW5n' +
        'dGg9PT0wKXtpbmZvLnRleHRDb250ZW50PSdObyByZXN1bHRzJzt9CiAgICBlbHNle2luZm8udGV4dENv' +
        'bnRlbnQ9J1Nob3dpbmcgJysoc3RhcnQrMSkrJ+KAkycrZW5kKycgb2YgJytfZmlsdGVyZWQubGVuZ3Ro' +
        'KycgZm9sZGVycyc7fQogIH0KCiAgLy8gVXBkYXRlIHByZXYvbmV4dAogIGRvY3VtZW50LmdldEVsZW1l' +
        'bnRCeUlkKCdidG5QcmV2JykuZGlzYWJsZWQ9KF9wYWdlPD0xKTsKICBkb2N1bWVudC5nZXRFbGVtZW50' +
        'QnlJZCgnYnRuTmV4dCcpLmRpc2FibGVkPShfcGFnZT49dG90YWwpOwoKICAvLyBVcGRhdGUgcGFnZSBu' +
        'dW1iZXIgYnV0dG9ucwogIHJlbmRlclBhZ2VCdXR0b25zKCk7CgogIC8vIEVtcHR5IHN0YXRlCiAgdmFy' +
        'IGVsPWRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCdlbXB0eVN0YXRlJyk7CiAgaWYoZWwpZWwuc3R5bGUu' +
        'ZGlzcGxheT0oX2ZpbHRlcmVkLmxlbmd0aD09PTApPydibG9jayc6J25vbmUnOwp9CgpmdW5jdGlvbiBy' +
        'ZW5kZXJQYWdlQnV0dG9ucygpewogIHZhciB0b3RhbD1NYXRoLmNlaWwoX2ZpbHRlcmVkLmxlbmd0aC9f' +
        'cGFnZVNpemUpfHwxOwogIHZhciBjb250YWluZXI9ZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoJ3BhZ2VC' +
        'dXR0b25zJyk7CiAgaWYoIWNvbnRhaW5lcilyZXR1cm47CiAgY29udGFpbmVyLmlubmVySFRNTD0nJzsK' +
        'CiAgdmFyIG1heEJ0bnM9NzsKICB2YXIgc3RhcnQ9TWF0aC5tYXgoMSxfcGFnZS1NYXRoLmZsb29yKG1h' +
        'eEJ0bnMvMikpOwogIHZhciBlbmQ9TWF0aC5taW4odG90YWwsc3RhcnQrbWF4QnRucy0xKTsKICBpZihl' +
        'bmQtc3RhcnQ8bWF4QnRucy0xKXN0YXJ0PU1hdGgubWF4KDEsZW5kLW1heEJ0bnMrMSk7CgogIGlmKHN0' +
        'YXJ0PjEpewogICAgdmFyIGI9bWFrZVBhZ2VCdG4oMSk7Y29udGFpbmVyLmFwcGVuZENoaWxkKGIpOwog' +
        'ICAgaWYoc3RhcnQ+Mil7dmFyIGU9ZG9jdW1lbnQuY3JlYXRlRWxlbWVudCgnc3BhbicpO2UudGV4dENv' +
        'bnRlbnQ9J+KApic7ZS5zdHlsZS5jc3NUZXh0PSdjb2xvcjp2YXIoLS10ZXh0Myk7cGFkZGluZzowIDRw' +
        'eDtmb250LXNpemU6MTFweCc7Y29udGFpbmVyLmFwcGVuZENoaWxkKGUpO30KICB9CiAgZm9yKHZhciBp' +
        'PXN0YXJ0O2k8PWVuZDtpKyspe2NvbnRhaW5lci5hcHBlbmRDaGlsZChtYWtlUGFnZUJ0bihpKSk7fQog' +
        'IGlmKGVuZDx0b3RhbCl7CiAgICBpZihlbmQ8dG90YWwtMSl7dmFyIGUyPWRvY3VtZW50LmNyZWF0ZUVs' +
        'ZW1lbnQoJ3NwYW4nKTtlMi50ZXh0Q29udGVudD0n4oCmJztlMi5zdHlsZS5jc3NUZXh0PSdjb2xvcjp2' +
        'YXIoLS10ZXh0Myk7cGFkZGluZzowIDRweDtmb250LXNpemU6MTFweCc7Y29udGFpbmVyLmFwcGVuZENo' +
        'aWxkKGUyKTt9CiAgICBjb250YWluZXIuYXBwZW5kQ2hpbGQobWFrZVBhZ2VCdG4odG90YWwpKTsKICB9' +
        'Cn0KCmZ1bmN0aW9uIG1ha2VQYWdlQnRuKG4pewogIHZhciBiPWRvY3VtZW50LmNyZWF0ZUVsZW1lbnQo' +
        'J2J1dHRvbicpOwogIGIudGV4dENvbnRlbnQ9bjsKICBiLmNsYXNzTmFtZT0ncGFnZS1idG4nKyhuPT09' +
        'X3BhZ2U/JyBhY3RpdmUnOicnKTsKICBiLm9uY2xpY2s9KGZ1bmN0aW9uKHBnKXtyZXR1cm4gZnVuY3Rp' +
        'b24oKXtzaG93UGFnZShwZyk7fTt9KShuKTsKICByZXR1cm4gYjsKfQoKZnVuY3Rpb24gZXhwYW5kQWxs' +
        'KCl7CiAgdmFyIHN0YXJ0PShfcGFnZS0xKSpfcGFnZVNpemU7CiAgdmFyIGVuZD1NYXRoLm1pbihzdGFy' +
        'dCtfcGFnZVNpemUsX2ZpbHRlcmVkLmxlbmd0aCk7CiAgZm9yKHZhciBpPXN0YXJ0O2k8ZW5kO2krKyl7' +
        'CiAgICB2YXIgZD1fZmlsdGVyZWRbaV0ucXVlcnlTZWxlY3RvcignLmZvbGRlci1kZXRhaWxzJyk7CiAg' +
        'ICB2YXIgaWM9X2ZpbHRlcmVkW2ldLnF1ZXJ5U2VsZWN0b3IoJy50b2dnbGUtaWNvbicpOwogICAgaWYo' +
        'ZCl7ZC5jbGFzc0xpc3QuYWRkKCdvcGVuJyk7fQogICAgaWYoaWMpe2ljLnN0eWxlLnRyYW5zZm9ybT0n' +
        'cm90YXRlKDkwZGVnKSc7fQogIH0KfQoKZnVuY3Rpb24gY29sbGFwc2VBbGwoKXsKICB2YXIgc3RhcnQ9' +
        'KF9wYWdlLTEpKl9wYWdlU2l6ZTsKICB2YXIgZW5kPU1hdGgubWluKHN0YXJ0K19wYWdlU2l6ZSxfZmls' +
        'dGVyZWQubGVuZ3RoKTsKICBmb3IodmFyIGk9c3RhcnQ7aTxlbmQ7aSsrKXsKICAgIHZhciBkPV9maWx0' +
        'ZXJlZFtpXS5xdWVyeVNlbGVjdG9yKCcuZm9sZGVyLWRldGFpbHMnKTsKICAgIHZhciBpYz1fZmlsdGVy' +
        'ZWRbaV0ucXVlcnlTZWxlY3RvcignLnRvZ2dsZS1pY29uJyk7CiAgICBpZihkKXtkLmNsYXNzTGlzdC5y' +
        'ZW1vdmUoJ29wZW4nKTt9CiAgICBpZihpYyl7aWMuc3R5bGUudHJhbnNmb3JtPScnO30KICB9Cn0KCmRv' +
        'Y3VtZW50LmFkZEV2ZW50TGlzdGVuZXIoJ0RPTUNvbnRlbnRMb2FkZWQnLGZ1bmN0aW9uKCl7aW5pdCgp' +
        'O30pOwo='
    )

    $css = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($cssB64))
    $jsB64decoded = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($jsB64))
    $js = $jsB64decoded

    # ── Build folder rows ─────────────────────────────────────
    $rows = [System.Text.StringBuilder]::new()

    foreach ($fd in $FolderData) {
        $path    = $fd.Path
        $relPath = $fd.Path.Substring([Math]::Min($RootPath.TrimEnd('\').Length + 1, $fd.Path.Length))
        if ($relPath -eq "") { $relPath = $fd.Path }

        $statusClass = if ($fd.Depth -eq 0)              { "status-root" }
                       elseif ($null -ne $fd.Error)       { "status-error" }
                       elseif ($fd.HasChanges)             { "status-changed" }
                       else                                { "status-inherited" }

        $statusLabel = if ($fd.Depth -eq 0)              { "Root" }
                       elseif ($null -ne $fd.Error)       { "Error" }
                       elseif ($fd.HasChanges)             { "Changed" }
                       else                                { "Inherited" }

        $inheritText  = if ($fd.InheritanceEnabled) { "Inherited" } else { "Protected" }
        $inheritClass = if ($fd.InheritanceEnabled) { "inherit-yes" } else { "inherit-no" }
        $ownerEnc     = [System.Web.HttpUtility]::HtmlEncode($fd.Owner)
        $pathEnc      = [System.Web.HttpUtility]::HtmlEncode($relPath)
        $rowId        = "row_$([Math]::Abs($path.GetHashCode()))"

        # Open/collapsed state - expanded by default only for root+changed
        $openClass = if ($fd.Depth -eq 0 -or $fd.HasChanges) { " open" } else { "" }
        $iconRot   = if ($fd.Depth -eq 0 -or $fd.HasChanges) { " style='transform:rotate(90deg)'" } else { "" }

        [void]$rows.Append("<div class='folder-row $statusClass'>")
        [void]$rows.Append("<div class='folder-header' onclick='toggle(""$rowId"",this)'>")
        [void]$rows.Append("<div class='folder-left'>")
        [void]$rows.Append("<span class='toggle-icon'$iconRot>&#9658;</span>")
        [void]$rows.Append("<span class='folder-icon'>&#128193;</span>")
        [void]$rows.Append("<span class='folder-path' title='$pathEnc'>$pathEnc</span>")
        [void]$rows.Append("</div>")
        [void]$rows.Append("<div class='folder-right'>")
        [void]$rows.Append("<span class='owner-badge'>$ownerEnc</span>")
        [void]$rows.Append("<span class='inherit-badge $inheritClass'>$inheritText</span>")
        [void]$rows.Append("<span class='status-badge $statusClass'>$statusLabel</span>")
        [void]$rows.Append("</div></div>")

        # ── Detail panel ─────────────────────────────────────
        [void]$rows.Append("<div class='folder-details$openClass' id='$rowId'>")

        if ($null -ne $fd.Error) {
            [void]$rows.Append("<div class='error-box'>&#10060; $([System.Web.HttpUtility]::HtmlEncode($fd.Error))</div>")
        } else {
            # Meta info row
            [void]$rows.Append("<div class='details-meta'>")
            [void]$rows.Append("<span><strong>Path:</strong> $([System.Web.HttpUtility]::HtmlEncode($path))</span>")
            [void]$rows.Append("<span><strong>Owner:</strong> $ownerEnc</span>")
            [void]$rows.Append("<span><strong>Inheritance:</strong> $inheritText</span>")
            [void]$rows.Append("</div>")

            $explicit  = @($fd.Rules | Where-Object { -not $_.IsInherited })
            $inherited = @($fd.Rules | Where-Object { $_.IsInherited })

            if ($fd.Depth -gt 0 -and -not $fd.HasChanges) {
                [void]$rows.Append("<div class='same-as-parent'>&#10003; Same as parent folder — no local changes</div>")
            }

            # Explicit rules
            if ($explicit.Count -gt 0) {
                [void]$rows.Append("<div class='rules-section'>")
                [void]$rows.Append("<div class='rules-section-title explicit-title'>Explicit Permissions</div>")
                [void]$rows.Append("<table class='rules-table'><thead><tr><th>Principal</th><th>Type</th><th>Permission</th><th>Applies to</th><th>Kind</th></tr></thead><tbody>")
                foreach ($rule in $explicit) {
                    $rowClass = if ($rule.AccessType -eq "Deny") { " class='deny-rule'" } else { "" }
                    $typeBadge = if ($rule.AccessType -eq "Deny") { "<span class='badge badge-deny'>Deny</span>" } else { "<span class='badge badge-allow'>Allow</span>" }
                    [void]$rows.Append("<tr$rowClass><td class='identity'>$([System.Web.HttpUtility]::HtmlEncode($rule.Identity))</td>")
                    [void]$rows.Append("<td>$typeBadge</td>")
                    [void]$rows.Append("<td>$([System.Web.HttpUtility]::HtmlEncode($rule.RightsSimple))</td>")
                    [void]$rows.Append("<td class='scope'>$([System.Web.HttpUtility]::HtmlEncode($rule.InheritanceDesc))</td>")
                    [void]$rows.Append("<td><span class='badge badge-explicit'>Explicit</span></td></tr>")
                }
                [void]$rows.Append("</tbody></table></div>")
            }

            # Inherited rules
            if ($inherited.Count -gt 0) {
                [void]$rows.Append("<div class='rules-section'>")
                [void]$rows.Append("<div class='rules-section-title inherited-title'>Inherited Permissions</div>")
                [void]$rows.Append("<table class='rules-table inherited-table'><thead><tr><th>Principal</th><th>Type</th><th>Permission</th><th>Applies to</th><th>Kind</th></tr></thead><tbody>")
                foreach ($rule in $inherited) {
                    $rowClass  = if ($rule.AccessType -eq "Deny") { " class='deny-rule'" } else { "" }
                    $typeBadge = if ($rule.AccessType -eq "Deny") { "<span class='badge badge-deny'>Deny</span>" } else { "<span class='badge badge-allow'>Allow</span>" }
                    [void]$rows.Append("<tr$rowClass><td class='identity'>$([System.Web.HttpUtility]::HtmlEncode($rule.Identity))</td>")
                    [void]$rows.Append("<td>$typeBadge</td>")
                    [void]$rows.Append("<td>$([System.Web.HttpUtility]::HtmlEncode($rule.RightsSimple))</td>")
                    [void]$rows.Append("<td class='scope'>$([System.Web.HttpUtility]::HtmlEncode($rule.InheritanceDesc))</td>")
                    [void]$rows.Append("<td><span class='badge badge-inherited'>Inherited</span></td></tr>")
                }
                [void]$rows.Append("</tbody></table></div>")
            }

            if ($explicit.Count -eq 0 -and $inherited.Count -eq 0) {
                [void]$rows.Append("<div class='no-rules'>No permission rules found.</div>")
            }
        }
        [void]$rows.Append("</div></div>")
    }

    $rootEnc = [System.Web.HttpUtility]::HtmlEncode($RootPath)
    $userEnc = [System.Web.HttpUtility]::HtmlEncode($currentUser)

    $out = [System.Text.StringBuilder]::new()
    [void]$out.Append("<!DOCTYPE html><html lang='en'><head><meta charset='UTF-8'><meta name='viewport' content='width=device-width,initial-scale=1'>")
    [void]$out.Append("<title>ACLens Report</title>")
    [void]$out.Append("<style>$css</style></head>")
    [void]$out.Append("<body>")

    # Header
    [void]$out.Append("<div class='report-header'>")
    [void]$out.Append("<div class='header-left'>")
    [void]$out.Append("<div class='report-title'><span class='icon'>&#128274;</span> ACLens &mdash; NTFS Permission Report</div>")
    [void]$out.Append("<div class='report-subtitle'>$rootEnc</div>")
    [void]$out.Append("</div>")
    [void]$out.Append("<div class='header-right'>")
    [void]$out.Append("<div class='report-subtitle'>$reportDate &bull; $computerName &bull; $userEnc</div>")
    [void]$out.Append("<button id='themeBtn' class='theme-toggle' onclick='toggleTheme()'>&#9728;&#65039; Light</button>")
    [void]$out.Append("</div></div>")

    # Stats bar
    [void]$out.Append("<div class='stats-bar'>")
    [void]$out.Append("<div class='stat-card stat-total'><span class='stat-number'>$totalFolders</span><span class='stat-label'>Total folders</span></div>")
    [void]$out.Append("<div class='stat-card stat-changed'><span class='stat-number'>$changedFolders</span><span class='stat-label'>Changed</span></div>")
    [void]$out.Append("<div class='stat-card stat-inherited'><span class='stat-number'>$inheritedOnly</span><span class='stat-label'>Inherited only</span></div>")
    [void]$out.Append("<div class='stat-card stat-error'><span class='stat-number'>$errorFolders</span><span class='stat-label'>Errors</span></div>")
    [void]$out.Append("</div>")

    # Agenda (collapsible legend)
    [void]$out.Append("<div class='agenda-bar'><button class='agenda-toggle' onclick='toggleAgenda(this)'>")
    [void]$out.Append("<span class='agenda-toggle-icon'>&#9658;</span> &#128218; Legend &amp; Reference ")
    [void]$out.Append("<span class='agenda-toggle-hint'>click to expand</span></button>")
    [void]$out.Append("<div class='agenda-content' id='agendaContent' style='display:none'>")
    [void]$out.Append("<div class='agenda-grid'>")
    [void]$out.Append("<div class='agenda-section'><div class='agenda-section-title'>Folder Status</div>")
    [void]$out.Append("<div class='agenda-row'><span class='agenda-dot' style='background:#A78BFA'></span><span class='agenda-label'>Root</span><span class='agenda-desc'>The start folder of the analysis</span></div>")
    [void]$out.Append("<div class='agenda-row'><span class='agenda-dot' style='background:#FBBF24'></span><span class='agenda-label'>Changed</span><span class='agenda-desc'>Permissions differ from parent folder</span></div>")
    [void]$out.Append("<div class='agenda-row'><span class='agenda-dot' style='background:#34D399'></span><span class='agenda-label'>Inherited</span><span class='agenda-desc'>All permissions passed down from parent</span></div>")
    [void]$out.Append("<div class='agenda-row'><span class='agenda-dot' style='background:#F87171'></span><span class='agenda-label'>Error</span><span class='agenda-desc'>Could not read permissions</span></div>")
    [void]$out.Append("</div>")
    [void]$out.Append("<div class='agenda-section'><div class='agenda-section-title'>Permission Types</div>")
    [void]$out.Append("<div class='agenda-row'><span class='agenda-perm'>Full Control</span><span class='agenda-desc'>Read, write, change permissions, take ownership</span></div>")
    [void]$out.Append("<div class='agenda-row'><span class='agenda-perm'>Modify</span><span class='agenda-desc'>Read, write, delete files and subfolders</span></div>")
    [void]$out.Append("<div class='agenda-row'><span class='agenda-perm'>Read &amp; Execute</span><span class='agenda-desc'>View and run files</span></div>")
    [void]$out.Append("<div class='agenda-row'><span class='agenda-perm'>Read</span><span class='agenda-desc'>View files and folder contents</span></div>")
    [void]$out.Append("<div class='agenda-row'><span class='agenda-perm'>Write</span><span class='agenda-desc'>Create files and subfolders</span></div>")
    [void]$out.Append("</div>")
    [void]$out.Append("<div class='agenda-section'><div class='agenda-section-title'>Rule Types</div>")
    [void]$out.Append("<div class='agenda-row'><span class='agenda-dot' style='background:#A78BFA'></span><span class='agenda-label'>Explicit</span><span class='agenda-desc'>Set directly on this folder</span></div>")
    [void]$out.Append("<div class='agenda-row'><span class='agenda-dot' style='background:#4B5563'></span><span class='agenda-label'>Inherited</span><span class='agenda-desc'>Passed down from a parent folder</span></div>")
    [void]$out.Append("<div class='agenda-row'><span class='agenda-dot' style='background:#34D399'></span><span class='agenda-label'>Allow</span><span class='agenda-desc'>Grants access — can be overridden by Deny</span></div>")
    [void]$out.Append("<div class='agenda-row'><span class='agenda-dot' style='background:#F87171'></span><span class='agenda-label'>Deny</span><span class='agenda-desc'>Blocks access — always overrides Allow</span></div>")
    [void]$out.Append("</div></div></div></div>")

    # Toolbar
    [void]$out.Append("<div class='toolbar'>")
    [void]$out.Append("<span class='filter-group'>")
    [void]$out.Append("<button class='btn filter-btn active' onclick='filterRows(""all"",this)'>All</button>")
    [void]$out.Append("<button class='btn filter-btn' onclick='filterRows(""changed"",this)'>Changed</button>")
    [void]$out.Append("<button class='btn filter-btn' onclick='filterRows(""inherited"",this)'>Inherited only</button>")
    [void]$out.Append("<button class='btn filter-btn' onclick='filterRows(""error"",this)'>Errors</button>")
    [void]$out.Append("</span>")
    [void]$out.Append("<input class='search-box' type='text' placeholder='Search path or user...' oninput='searchRows(this.value)'>")
    [void]$out.Append("<div class='toolbar-spacer'></div>")
    [void]$out.Append("<button class='btn' onclick='expandAll()'>Expand page</button>")
    [void]$out.Append("<button class='btn' onclick='collapseAll()'>Collapse page</button>")
    [void]$out.Append("<button class='btn' onclick='window.print()'>&#128438; Print</button>")
    [void]$out.Append("</div>")

    # Page nav
    [void]$out.Append("<div class='page-nav'>")
    [void]$out.Append("<button id='btnPrev' class='page-btn' onclick='showPage(_page-1)'>&#8592;</button>")
    [void]$out.Append("<span id='pageButtons' style='display:flex;align-items:center;gap:3px'></span>")
    [void]$out.Append("<button id='btnNext' class='page-btn' onclick='showPage(_page+1)'>&#8594;</button>")
    [void]$out.Append("<select class='page-size-select' onchange='setPageSize(this.value)'>")
    [void]$out.Append("<option value='25'>25 / page</option>")
    [void]$out.Append("<option value='50'>50 / page</option>")
    [void]$out.Append("<option value='100'>100 / page</option>")
    [void]$out.Append("<option value='9999'>All</option>")
    [void]$out.Append("</select>")
    [void]$out.Append("<span id='pageInfo' class='page-info'></span>")
    [void]$out.Append("</div>")

    # Folder list
    [void]$out.Append("<div class='folder-list'>")
    [void]$out.Append($rows.ToString())
    [void]$out.Append("<div id='emptyState' class='empty-page' style='display:none'><div class='icon'>&#128269;</div><p>No folders match this filter.</p></div>")
    [void]$out.Append("</div>")

    # Footer
    [void]$out.Append("<div class='report-footer'>ACLens v0.4.0-beta &bull; $reportDate &bull; $computerName &bull; $userEnc</div>")
    [void]$out.Append("<script>$js</script>")
    [void]$out.Append("</body></html>")

    [System.IO.File]::WriteAllText($OutputPath, $out.ToString(), [System.Text.UTF8Encoding]::new($true))
}


function New-CompareHTMLReport {
    param([hashtable]$OldData, [hashtable]$NewData, [string]$OldLabel, [string]$NewLabel, [string]$OutputPath)

    $reportDate = Get-Date -Format "dd.MM.yyyy HH:mm:ss"

    $allPaths = @($OldData.Keys) + @($NewData.Keys) | Sort-Object -Unique

    $addedFolders   = @()
    $removedFolders = @()
    $changedFolders = @()
    $rows = [System.Text.StringBuilder]::new()

    foreach ($path in $allPaths) {
        $inOld = $OldData.ContainsKey($path)
        $inNew = $NewData.ContainsKey($path)

        if (-not $inOld -and $inNew) {
            $addedFolders += $path
            [void]$rows.Append("<tr class='row-added'>")
            [void]$rows.Append("<td class='path-cell'>&#43; $([System.Web.HttpUtility]::HtmlEncode($path))</td>")
            [void]$rows.Append("<td><span class='diff-badge badge-added'>Added</span></td>")
            [void]$rows.Append("<td colspan='2' class='diff-note'>New folder &mdash; not present in baseline</td>")
            [void]$rows.Append("</tr>")
        }
        elseif ($inOld -and -not $inNew) {
            $removedFolders += $path
            [void]$rows.Append("<tr class='row-removed'>")
            [void]$rows.Append("<td class='path-cell'>&#8722; $([System.Web.HttpUtility]::HtmlEncode($path))</td>")
            [void]$rows.Append("<td><span class='diff-badge badge-removed'>Removed</span></td>")
            [void]$rows.Append("<td colspan='2' class='diff-note'>Folder removed &mdash; not present in current scan</td>")
            [void]$rows.Append("</tr>")
        }
        else {
            $old = $OldData[$path]
            $new = $NewData[$path]
            $oldKey = Get-RulesKey $old.Rules
            $newKey = Get-RulesKey $new.Rules
            $sameInherit = ($old.InheritanceEnabled -eq $new.InheritanceEnabled)

            if ($oldKey -ne $newKey -or -not $sameInherit -or $old.Owner -ne $new.Owner) {
                $changedFolders += $path

                # Find added/removed rules
                $oldRules = @{}; foreach ($r in $old.Rules) { $oldRules["$($r.Identity)|$($r.AccessType)|$($r.Rights)|$($r.InheritanceFlags)"] = $r }
                $newRules = @{}; foreach ($r in $new.Rules) { $newRules["$($r.Identity)|$($r.AccessType)|$($r.Rights)|$($r.InheritanceFlags)"] = $r }

                $addedRules   = $newRules.Keys | Where-Object { -not $oldRules.ContainsKey($_) }
                $removedRules = $oldRules.Keys | Where-Object { -not $newRules.ContainsKey($_) }

                [void]$rows.Append("<tr class='row-changed'>")
                [void]$rows.Append("<td class='path-cell' rowspan='999'>&#9998; $([System.Web.HttpUtility]::HtmlEncode($path))</td>")
                [void]$rows.Append("<td><span class='diff-badge badge-changed'>Changed</span></td>")

                $details = [System.Text.StringBuilder]::new()

                if ($old.Owner -ne $new.Owner) {
                    [void]$details.Append("<div class='diff-detail'><span class='diff-key'>Owner:</span> <span class='old-val'>$([System.Web.HttpUtility]::HtmlEncode($old.Owner))</span> &rarr; <span class='new-val'>$([System.Web.HttpUtility]::HtmlEncode($new.Owner))</span></div>")
                }
                if (-not $sameInherit) {
                    $oldI = if ($old.InheritanceEnabled) { 'Active' } else { 'Disabled' }
                    $newI = if ($new.InheritanceEnabled) { 'Active' } else { 'Disabled' }
                    [void]$details.Append("<div class='diff-detail'><span class='diff-key'>Inheritance:</span> <span class='old-val'>$oldI</span> &rarr; <span class='new-val'>$newI</span></div>")
                }
                foreach ($k in $addedRules) {
                    $r = $newRules[$k]
                    [void]$details.Append("<div class='diff-detail added-rule'>&#43; $([System.Web.HttpUtility]::HtmlEncode($r.Identity)) &mdash; $($r.AccessType) &mdash; $($r.Rights)</div>")
                }
                foreach ($k in $removedRules) {
                    $r = $oldRules[$k]
                    [void]$details.Append("<div class='diff-detail removed-rule'>&#8722; $([System.Web.HttpUtility]::HtmlEncode($r.Identity)) &mdash; $($r.AccessType) &mdash; $($r.Rights)</div>")
                }

                [void]$rows.Append("<td colspan='2'>$($details.ToString())</td></tr>")
            }
        }
    }

    $totalChanged = $changedFolders.Count
    $totalAdded   = $addedFolders.Count
    $totalRemoved = $removedFolders.Count
    $totalSame    = $allPaths.Count - $totalChanged - $totalAdded - $totalRemoved

    $cssCompare = ':root{--bg:#1E1E2E;--bg2:#2D2D3F;--bg3:#1A1A2E;--bg4:#252535;--border:#4B5563;--border2:#374151;--text:#E2E8F0;--text2:#9CA3AF;--text3:#6B7280;--accent:#A78BFA;--green:#34D399;--red:#F87171;--yellow:#FBBF24;--font-mono:Consolas,monospace;--font-ui:"Segoe UI",system-ui,sans-serif;--radius:6px}*{box-sizing:border-box;margin:0;padding:0}body{background:var(--bg);color:var(--text);font-family:var(--font-ui);font-size:13px;line-height:1.5}.header{background:linear-gradient(180deg,#2D2D3F,#1E1E2E);border-bottom:1px solid var(--border);padding:24px 32px 18px}.title{font-size:20px;font-weight:700;color:var(--accent);margin-bottom:6px}.subtitle{color:var(--text2);font-size:12px;margin-bottom:12px}.meta{display:flex;gap:24px;flex-wrap:wrap}.meta-item{display:flex;flex-direction:column;gap:2px}.meta-lbl{font-size:10px;color:var(--text3);text-transform:uppercase;letter-spacing:.8px}.meta-val{font-size:13px;font-weight:500}.stats{display:flex;gap:10px;padding:14px 32px;background:var(--bg2);border-bottom:1px solid var(--border);flex-wrap:wrap}.stat{background:var(--bg4);border:1px solid var(--border);border-radius:var(--radius);padding:9px 16px;display:flex;flex-direction:column;align-items:center;min-width:110px}.stat-num{font-size:22px;font-weight:700;line-height:1}.stat-lbl{font-size:11px;color:var(--text2);margin-top:2px}.s-changed .stat-num{color:var(--yellow)}.s-added .stat-num{color:var(--green)}.s-removed .stat-num{color:var(--red)}.s-same .stat-num{color:var(--text2)}.table-wrap{padding:16px 24px;overflow-x:auto}table{width:100%;border-collapse:collapse;font-size:12px}thead tr{background:var(--bg4)}th{text-align:left;padding:7px 10px;color:var(--text2);font-weight:500;font-size:11px;border-bottom:1px solid var(--border)}td{padding:7px 10px;border-bottom:1px solid var(--border2);vertical-align:top}.path-cell{font-family:var(--font-mono);font-size:11px;color:var(--text);min-width:260px}.row-added{background:rgba(52,211,153,.06)}.row-added .path-cell{color:var(--green)}.row-removed{background:rgba(241,113,113,.06)}.row-removed .path-cell{color:var(--red)}.row-changed{background:rgba(251,191,36,.05)}.row-changed .path-cell{color:var(--yellow)}.diff-badge{display:inline-block;font-size:10px;padding:2px 8px;border-radius:10px;font-weight:600;white-space:nowrap}.badge-added{background:rgba(52,211,153,.15);border:1px solid rgba(52,211,153,.4);color:var(--green)}.badge-removed{background:rgba(241,113,113,.15);border:1px solid rgba(241,113,113,.4);color:var(--red)}.badge-changed{background:rgba(251,191,36,.15);border:1px solid rgba(251,191,36,.4);color:var(--yellow)}.diff-note{color:var(--text2)}.diff-detail{margin:2px 0;font-size:11px;color:var(--text2)}.diff-key{font-weight:600;color:var(--text);margin-right:4px}.old-val{color:var(--red);text-decoration:line-through;margin-right:4px}.new-val{color:var(--green)}.added-rule{color:var(--green)}.removed-rule{color:var(--red)}.footer{text-align:center;padding:16px;color:var(--text3);font-size:11px;border-top:1px solid var(--border2);background:var(--bg3)}.no-diff{padding:40px 32px;text-align:center;color:var(--text3);font-size:14px}'

    $noRows = ($totalChanged + $totalAdded + $totalRemoved) -eq 0

    $html  = "<!DOCTYPE html><html lang='en'><head><meta charset='UTF-8'><title>ACLens Compare</title>"
    $html += "<style>$cssCompare</style></head><body>"
    $html += "<div class='header'>"
    $html += "<div class='title'>&#128270; ACLens &mdash; Permission Diff</div>"
    $html += "<div class='subtitle'>Baseline vs. Current Scan</div>"
    $html += "<div class='meta'>"
    $html += "<div class='meta-item'><span class='meta-lbl'>Baseline</span><span class='meta-val'>$([System.Web.HttpUtility]::HtmlEncode($OldLabel))</span></div>"
    $html += "<div class='meta-item'><span class='meta-lbl'>Current Scan</span><span class='meta-val'>$([System.Web.HttpUtility]::HtmlEncode($NewLabel))</span></div>"
    $html += "<div class='meta-item'><span class='meta-lbl'>Generated</span><span class='meta-val'>$reportDate</span></div>"
    $html += "</div></div>"
    $html += "<div class='stats'>"
    $html += "<div class='stat s-changed'><span class='stat-num'>$totalChanged</span><span class='stat-lbl'>Permissions changed</span></div>"
    $html += "<div class='stat s-added'><span class='stat-num'>$totalAdded</span><span class='stat-lbl'>Folders added</span></div>"
    $html += "<div class='stat s-removed'><span class='stat-num'>$totalRemoved</span><span class='stat-lbl'>Folders removed</span></div>"
    $html += "<div class='stat s-same'><span class='stat-num'>$totalSame</span><span class='stat-lbl'>Unchanged</span></div>"
    $html += "</div>"

    if ($noRows) {
        $html += "<div class='no-diff'>&#10003; No differences found &mdash; permissions are identical.</div>"
    } else {
        $html += "<div class='table-wrap'><table>"
        $html += "<thead><tr><th>Path</th><th>Status</th><th colspan='2'>Details</th></tr></thead><tbody>"
        $html += $rows.ToString()
        $html += "</tbody></table></div>"
    }

    $html += "<div class='footer'>ACLens Compare &bull; $reportDate</div>"
    $html += "</body></html>"

    [System.IO.File]::WriteAllText($OutputPath, $html, [System.Text.UTF8Encoding]::new($true))
    return [PSCustomObject]@{ Changed=$totalChanged; Added=$totalAdded; Removed=$totalRemoved }
}



# ============================================================
# MODULE: GUI
# ============================================================
# ACLens - GUI.psm1
# Main application window and all UI logic

function Start-ACLensGUI {

    [System.Windows.Forms.Application]::EnableVisualStyles()

# Farben
$C_BG       = HC '#1E1E2E'
$C_INPUT    = HC '#2D2D3F'
$C_TITLE    = HC '#1A1A2E'
$C_FOOTER   = HC '#111827'
$C_BORDER   = HC '#4B5563'
$C_BORD2    = HC '#374151'
$C_BTN      = HC '#374151'
$C_PRIMARY  = HC '#7C3AED'
$C_PRIM2    = HC '#4C1D95'
$C_CANCEL   = HC '#DC2626'
$C_CANC2    = HC '#B91C1C'
$C_TXT      = HC '#E2E8F0'
$C_ACCENT   = HC '#A78BFA'
$C_MID      = HC '#9CA3AF'
$C_LOW      = HC '#6B7280'
$C_OK       = HC '#34D399'
$C_ERR      = HC '#F87171'
$C_WARN     = HC '#FBBF24'

# Fonts
$FN = [System.Drawing.Font]::new("Segoe UI", 9)
$FB = [System.Drawing.Font]::new("Segoe UI", 9,  [System.Drawing.FontStyle]::Bold)
$FT = [System.Drawing.Font]::new("Segoe UI", 15, [System.Drawing.FontStyle]::Bold)
$FS = [System.Drawing.Font]::new("Segoe UI", 8)

# Konstanten
$PAD    = 20   # Seitenabstand
$HDR_H  = 54   # Header-Höhe
$FTR_H  = 62   # Footer-Höhe
$ROW_H  = 54   # Höhe pro Eingabezeile (Label + Feld)
$BTN_W  = 110  # Breite Browse/Save Buttons
$BTN_H  = 26   # Höhe Browse/Save Buttons

# ── Form ─────────────────────────────────────────────────────
$form = [System.Windows.Forms.Form]::new()
$form.Text            = "ACLens"
$form.Size            = [System.Drawing.Size]::new(760, 480)
$form.MinimumSize     = [System.Drawing.Size]::new(820, 440)
$form.StartPosition   = "CenterScreen"
$form.FormBorderStyle = "Sizable"
$form.MaximizeBox     = $true
$form.BackColor       = $C_BG
$form.ForeColor       = HC "#E2E8F0"
$form.Font            = $FN

# ── Header ───────────────────────────────────────────────────
$pnlHdr = [System.Windows.Forms.Panel]::new()
$pnlHdr.SetBounds(0, 0, $form.ClientSize.Width, $HDR_H)
$pnlHdr.BackColor = $C_TITLE
$pnlHdr.Anchor    = "Top,Left,Right"
$form.Controls.Add($pnlHdr)

$lblTitle = [System.Windows.Forms.Label]::new()
$lblTitle.Text      = "ACLens"
$lblTitle.Font      = $FT
$lblTitle.ForeColor = $C_ACCENT
$lblTitle.BackColor = [System.Drawing.Color]::Transparent
$lblTitle.AutoSize  = $true
$lblTitle.Location  = [System.Drawing.Point]::new(18, 10)
$pnlHdr.Controls.Add($lblTitle)

# Version badge next to title
# Version badge — same style as title, inline
$lblVersion = [System.Windows.Forms.Label]::new()
$lblVersion.Text      = "v0.4.0-beta"
$lblVersion.Font      = [System.Drawing.Font]::new("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$lblVersion.ForeColor = HC '#A78BFA'
$lblVersion.BackColor = [System.Drawing.Color]::Transparent
$lblVersion.AutoSize  = $true
$lblVersion.Location  = [System.Drawing.Point]::new(100, 14)   # will be adjusted after lblTitle loads
$pnlHdr.Controls.Add($lblVersion)

# ── Light mode toggle button ──────────────────────────────────
$script:lightMode = $false
$btnTheme = [System.Windows.Forms.Button]::new()
$btnTheme.Text      = $script:STR.BtnThemeLight
$btnTheme.Size      = [System.Drawing.Size]::new(72, 24)
$btnTheme.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnTheme.Font      = $FS
$btnTheme.Cursor    = "Hand"
$btnTheme.BackColor = HC "#2D2D3F"
$btnTheme.ForeColor = HC "#9CA3AF"
$btnTheme.FlatAppearance.BorderSize  = 1
    $btnTheme.UseVisualStyleBackColor = $false
$btnTheme.FlatAppearance.BorderColor = HC "#4B5563"
$pnlHdr.Controls.Add($btnTheme)

function Apply-Theme {
    if ($script:lightMode) {
        $form.BackColor       = HC "#F5F5F5"
        $pnlHdr.BackColor     = HC "#ECECEC"
        $pnlTabs.BackColor    = HC "#E2E2E2"
        $tabLine.BackColor    = HC "#7C3AED"
        $tabSep.BackColor     = HC "#CCCCCC"
        $pnlIn.BackColor      = HC "#F5F5F5"
        $pnlSP.BackColor      = HC "#F5F5F5"
        $pnlFtr.BackColor     = HC "#E2E2E2"
        $sepFtr.BackColor     = HC "#BBBBBB"
        $sepHdr.BackColor     = HC "#BBBBBB"
        $sepMid.BackColor     = HC "#CCCCCC"
        $lblTitle.ForeColor   = HC "#7C3AED"
        $lblVersion.ForeColor = HC "#7C3AED"
        $lblVersion.BackColor = HC "#EDE9FE"
        $lblSub.ForeColor     = HC "#555555"
        $lblStatus.ForeColor  = HC "#333333"
        $lblStats.ForeColor   = HC "#333333"
        $tbPath.BackColor     = HC "#FFFFFF"
        $tbPath.ForeColor     = HC "#111111"
        $tbOut.BackColor      = HC "#FFFFFF"
        $tbOut.ForeColor      = HC "#555555"
        $numDepth.BackColor   = HC "#FFFFFF"
        $numDepth.ForeColor   = HC "#111111"
        $lblPath.ForeColor    = HC "#111111"
        $lblOut.ForeColor     = HC "#111111"
        $lblDepth.ForeColor   = HC "#111111"
        $lblUnlim.ForeColor   = HC "#555555"
        $chk.ForeColor        = HC "#222222"
        $btnTheme.Text        = $script:STR.BtnThemeDark
        $btnTheme.BackColor   = HC "#E0E0E0"
        $btnTheme.ForeColor   = HC "#111111"
        $btnTheme.FlatAppearance.BorderColor = HC "#AAAAAA"
        # Footer buttons - keep colors, ensure white text
        $btnStart.ForeColor    = [System.Drawing.Color]::White
        $btnCancel.ForeColor   = [System.Drawing.Color]::White
        $btnOpenLast.ForeColor = [System.Drawing.Color]::White
$btnOpenLast.Font      = $FBTN
        $btnChangelog.ForeColor= [System.Drawing.Color]::White
        $btnCompare.ForeColor  = [System.Drawing.Color]::White
        # SP tab controls
        $spLblConnFg1 = if ($spLblConn.Text -eq "Connected") { HC "#059669" } else { HC "#DC2626" }
        $spLblConn.ForeColor = $spLblConnFg1
        $spLblConnDetail.ForeColor = HC "#444444"
        $spLblUrl.ForeColor        = HC "#111111"
        $spTbUrl.BackColor         = HC "#FFFFFF"
        $spTbUrl.ForeColor         = HC "#111111"
        $spLblDepth.ForeColor      = HC "#111111"
        $spNumDepth.BackColor      = HC "#FFFFFF"
        $spNumDepth.ForeColor      = HC "#111111"
        $spLblUnlim.ForeColor      = HC "#555555"
        $spChk.ForeColor           = HC "#222222"
    } else {
        $form.BackColor       = HC "#1E1E2E"
        $pnlHdr.BackColor     = HC "#1A1A2E"
        $pnlTabs.BackColor    = HC "#1A1A2E"
        $tabLine.BackColor    = HC "#7C3AED"
        $tabSep.BackColor     = HC "#4B5563"
        $pnlIn.BackColor      = HC "#1E1E2E"
        $pnlSP.BackColor      = HC "#1E1E2E"
        $pnlFtr.BackColor     = HC "#111827"
        $sepFtr.BackColor     = HC "#4B5563"
        $sepHdr.BackColor     = HC "#4B5563"
        $sepMid.BackColor     = HC "#4B5563"
        $lblTitle.ForeColor   = HC "#A78BFA"
        $lblVersion.ForeColor = HC "#A78BFA"
        $lblVersion.BackColor = HC "#2D1B69"
        $lblSub.ForeColor     = HC "#6B7280"
        $lblStatus.ForeColor  = HC "#9CA3AF"
        $lblStats.ForeColor   = HC "#9CA3AF"
        $btnBrowse.BackColor  = HC "#374151"
        $btnBrowse.ForeColor  = HC "#E2E8F0"
        $btnSave.BackColor    = HC "#374151"
        $btnSave.ForeColor    = HC "#E2E8F0"
        $btnSPSetup.BackColor = HC "#7C3AED"
        $btnSPSetup.ForeColor = [System.Drawing.Color]::White
        $tbPath.BackColor     = HC "#2D2D3F"
        $tbPath.ForeColor     = HC "#E2E8F0"
        $tbOut.BackColor      = HC "#2D2D3F"
        $tbOut.ForeColor      = HC "#6B7280"
        $numDepth.BackColor   = HC "#2D2D3F"
        $numDepth.ForeColor   = HC "#E2E8F0"
        $lblPath.ForeColor    = HC "#E2E8F0"
        $lblOut.ForeColor     = HC "#E2E8F0"
        $lblDepth.ForeColor   = HC "#E2E8F0"
        $lblUnlim.ForeColor   = HC "#6B7280"
        $chk.ForeColor        = HC "#9CA3AF"
        $btnTheme.Text        = $script:STR.BtnThemeLight
        $btnTheme.BackColor   = HC "#2D2D3F"
        $btnTheme.ForeColor   = HC "#9CA3AF"
        $btnTheme.FlatAppearance.BorderColor = HC "#4B5563"
        $btnStart.ForeColor    = [System.Drawing.Color]::White
        $btnCancel.ForeColor   = [System.Drawing.Color]::White
        $btnOpenLast.ForeColor = [System.Drawing.Color]::White
$btnOpenLast.Font      = $FBTN
        $btnChangelog.ForeColor= [System.Drawing.Color]::White
        $btnCompare.ForeColor  = [System.Drawing.Color]::White
        $spLblConnFg2 = if ($spLblConn.Text -eq "Connected") { HC "#34D399" } else { HC "#F87171" }
        $spLblConn.ForeColor = $spLblConnFg2
        $spLblConnDetail.ForeColor = HC "#9CA3AF"
        $spLblUrl.ForeColor        = HC "#E2E8F0"
        $spTbUrl.BackColor         = HC "#2D2D3F"
        $spTbUrl.ForeColor         = HC "#6B7280"
        $spLblDepth.ForeColor      = HC "#E2E8F0"
        $spNumDepth.BackColor      = HC "#2D2D3F"
        $spNumDepth.ForeColor      = HC "#E2E8F0"
        $spLblUnlim.ForeColor      = HC "#6B7280"
        $spChk.ForeColor           = HC "#9CA3AF"
    }
    # Re-apply active tab colors
    Set-ActiveTab $script:activeTab
}

$btnTheme.add_Click({
    $script:lightMode = -not $script:lightMode
    Apply-Theme
})

# Language toggle button
$btnLang = [System.Windows.Forms.Button]::new()
$btnLang.Text      = "DE"
$btnLang.Size      = [System.Drawing.Size]::new(38, 24)
$btnLang.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnLang.Font      = $FS
$btnLang.Cursor    = "Hand"
$btnLang.BackColor = HC "#2D2D3F"
$btnLang.ForeColor = HC "#9CA3AF"
$btnLang.FlatAppearance.BorderSize  = 1
    $btnLang.UseVisualStyleBackColor = $false
$btnLang.FlatAppearance.BorderColor = HC "#4B5563"
$pnlHdr.Controls.Add($btnLang)

function Apply-Language {
    $lblSub.Text        = $script:STR.AppSubtitle
    $lblPath.Text       = $script:STR.LblPath
    $lblOut.Text        = $script:STR.LblOutput
    $lblDepth.Text      = $script:STR.LblDepth
    $lblUnlim.Text      = $script:STR.LblUnlimited
    $btnBrowse.Text     = $script:STR.LblBrowse
    $btnSave.Text       = $script:STR.LblSaveAs
    $chk.Text           = $script:STR.ChkBrowser
    $lblStatus.Text     = $script:STR.LblReady
    $btnStart.Text      = $script:STR.BtnStart
    $btnCancel.Text     = $script:STR.BtnCancel
    $btnOpenLast.Text   = $script:STR.BtnOpenLast
    $btnChangelog.Text  = $script:STR.BtnChangelog
    $btnCompare.Text    = $script:STR.BtnCompare
    $spLblUrl.Text      = $script:STR.SpLblUrl
    $spLblDepth.Text    = $script:STR.SpLblDepth
    $spLblUnlim.Text    = $script:STR.LblUnlimited
    $spChk.Text         = $script:STR.SpChkBrowser
    $btnSPSetup.Text    = $script:STR.SpBtnSetup
    $tabNTFS.Text       = $script:STR.TabNTFS
    $tabSP.Text         = $script:STR.TabSP
    $btnLang.Text       = if ($script:lang -eq "EN") { "DE" } else { "EN" }
    if ($script:lightMode) {
        $btnTheme.Text  = $script:STR.BtnThemeDark
    } else {
        $btnTheme.Text  = $script:STR.BtnThemeLight
    }
    Update-SPStatus
}

$btnLang.add_Click({
    if ($script:lang -eq "EN") { Set-Language "DE" } else { Set-Language "EN" }
    Apply-Language
})

# Subtitle — right of version badge
$lblSub = [System.Windows.Forms.Label]::new()
$lblSub.Text      = $script:STR.AppSubtitle
$lblSub.Font      = $FS
$lblSub.ForeColor = HC '#6B7280'
$lblSub.BackColor = [System.Drawing.Color]::Transparent
$lblSub.AutoSize  = $true
$lblSub.Location  = [System.Drawing.Point]::new(200, 21)   # will be positioned after version label
$pnlHdr.Controls.Add($lblSub)

$sepHdr = [System.Windows.Forms.Panel]::new()
$sepHdr.SetBounds(0, $HDR_H, $form.ClientSize.Width, 1)
$sepHdr.BackColor = $C_BORDER
$sepHdr.Anchor    = "Top,Left,Right"
$form.Controls.Add($sepHdr)

# ── Footer ───────────────────────────────────────────────────
$pnlFtr = [System.Windows.Forms.Panel]::new()
$pnlFtr.BackColor = $C_FOOTER
$pnlFtr.ForeColor = [System.Drawing.Color]::White
$pnlFtr.Anchor    = "Bottom,Left,Right"
$form.Controls.Add($pnlFtr)

$sepFtr = [System.Windows.Forms.Panel]::new()
$sepFtr.BackColor = $C_BORDER
$sepFtr.Anchor    = "Bottom,Left,Right"
$form.Controls.Add($sepFtr)

$btnStart = [System.Windows.Forms.Button]::new()
$btnStart.Text      = $script:STR.BtnStart
$btnStart.Size      = [System.Drawing.Size]::new(160, 36)
$btnStart.Location  = [System.Drawing.Point]::new(16, 13)
$btnStart.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnStart.Font      = $FBTN
$btnStart.BackColor = $C_PRIMARY
$btnStart.ForeColor = [System.Drawing.Color]::White
$btnStart.UseVisualStyleBackColor = $false
$btnStart.Cursor    = "Hand"
$btnStart.FlatAppearance.BorderColor = $C_PRIM2
$btnStart.FlatAppearance.BorderSize  = 1
$pnlFtr.Controls.Add($btnStart)

$btnCancel = [System.Windows.Forms.Button]::new()
$btnCancel.Text      = $script:STR.BtnCancel
$btnCancel.Size      = [System.Drawing.Size]::new(96, 36)
$btnCancel.Location  = [System.Drawing.Point]::new(168, 13)
$btnCancel.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnCancel.Font      = $FBTN
$btnCancel.BackColor = $C_CANCEL
$btnCancel.ForeColor = [System.Drawing.Color]::White
$btnCancel.UseVisualStyleBackColor = $false
$btnCancel.Cursor    = "Hand"
$btnCancel.FlatAppearance.BorderColor = $C_CANC2
$btnCancel.FlatAppearance.BorderSize  = 1
$pnlFtr.Controls.Add($btnCancel)

$btnOpenLast = [System.Windows.Forms.Button]::new()
$btnOpenLast.Text      = $script:STR.BtnOpenLast
$btnOpenLast.Size      = [System.Drawing.Size]::new(162, 36)
$btnOpenLast.Location  = [System.Drawing.Point]::new(272, 13)
$btnOpenLast.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnOpenLast.Font      = $FBTN
$btnOpenLast.BackColor = HC '#2563EB'
$btnOpenLast.ForeColor = [System.Drawing.Color]::White
$btnOpenLast.UseVisualStyleBackColor = $false
$btnOpenLast.Font      = $FBTN
$btnOpenLast.Cursor    = "Hand"
$btnOpenLast.Enabled   = $false
$btnOpenLast.FlatAppearance.BorderColor = HC '#1E40AF'
$btnOpenLast.FlatAppearance.BorderSize  = 1
$pnlFtr.Controls.Add($btnOpenLast)

$btnChangelog = [System.Windows.Forms.Button]::new()
$btnChangelog.Text      = $script:STR.BtnChangelog
$btnChangelog.Size      = [System.Drawing.Size]::new(140, 36)
$btnChangelog.Location  = [System.Drawing.Point]::new(442, 13)
$btnChangelog.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnChangelog.Font      = $FBTN
$btnChangelog.ForeColor = [System.Drawing.Color]::White
$btnChangelog.UseVisualStyleBackColor = $false
$btnChangelog.BackColor = HC '#7C3AED'
$btnChangelog.ForeColor = [System.Drawing.Color]::White
$btnChangelog.Font      = $FBTN
$btnChangelog.Cursor    = "Hand"
$btnChangelog.FlatAppearance.BorderColor = HC '#4C1D95'
$btnChangelog.FlatAppearance.BorderSize  = 1
$pnlFtr.Controls.Add($btnChangelog)

$btnCompare = [System.Windows.Forms.Button]::new()
$btnCompare.Text      = $script:STR.BtnCompare
$btnCompare.Size      = [System.Drawing.Size]::new(148, 36)
$btnCompare.Location  = [System.Drawing.Point]::new(590, 13)
$btnCompare.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnCompare.Font      = $FBTN
$btnCompare.ForeColor = [System.Drawing.Color]::White
$btnCompare.UseVisualStyleBackColor = $false
$btnCompare.BackColor = HC '#2563EB'
$btnCompare.ForeColor = [System.Drawing.Color]::White
$btnCompare.Font      = $FBTN
$btnCompare.Cursor    = "Hand"
$btnCompare.FlatAppearance.BorderColor = HC '#1E40AF'
$btnCompare.FlatAppearance.BorderSize  = 1
$pnlFtr.Controls.Add($btnCompare)



# ── Tab bar (below header, above content) ────────────────────
$pnlTabs = [System.Windows.Forms.Panel]::new()
$pnlTabs.BackColor = HC '#1A1A2E'
$form.Controls.Add($pnlTabs)

$tabNTFS = [System.Windows.Forms.Button]::new()
$tabNTFS.Text = "  NTFS"; $tabNTFS.Size = [System.Drawing.Size]::new(120, 36)
$tabNTFS.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $tabNTFS.Font = $FB; $tabNTFS.Cursor = "Hand"
$tabNTFS.BackColor = HC '#1E1E2E'; $tabNTFS.ForeColor = HC '#A78BFA'; $tabNTFS.FlatAppearance.BorderSize = 0
$tabNTFS.FlatAppearance.BorderSize = 0
$pnlTabs.Controls.Add($tabNTFS)

$tabSP = [System.Windows.Forms.Button]::new()
$tabSP.Text = "  SharePoint"; $tabSP.Size = [System.Drawing.Size]::new(140, 36)
$tabSP.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $tabSP.Font = $FB; $tabSP.Cursor = "Hand"
$tabSP.BackColor = HC '#1A1A2E'; $tabSP.ForeColor = HC '#6B7280'
$tabSP.FlatAppearance.BorderSize = 0
$pnlTabs.Controls.Add($tabSP)

$tabLine = [System.Windows.Forms.Panel]::new()
$tabLine.BackColor = HC '#7C3AED'; $tabLine.Height = 2
$pnlTabs.Controls.Add($tabLine)

$tabSep = [System.Windows.Forms.Panel]::new()
$tabSep.BackColor = HC '#4B5563'; $tabSep.Height = 1
$form.Controls.Add($tabSep)

$script:activeTab = "NTFS"

# ── NTFS content panel ────────────────────────────────────────
$pnlIn = [System.Windows.Forms.Panel]::new()
$pnlIn.BackColor = HC '#1E1E2E'
$form.Controls.Add($pnlIn)

# ── SharePoint content panel ──────────────────────────────────
$pnlSP = [System.Windows.Forms.Panel]::new()
$pnlSP.BackColor = HC '#1E1E2E'
$pnlSP.Visible   = $false
$form.Controls.Add($pnlSP)

function Set-ActiveTab($tab) {
    $script:activeTab = $tab
    if ($script:lightMode) {
        $aBg = HC '#F5F5F5'; $aFg = HC '#7C3AED'
        $iBg = HC '#E2E2E2'; $iFg = HC '#555555'
    } else {
        $aBg = HC '#1E1E2E'; $aFg = HC '#A78BFA'
        $iBg = HC '#1A1A2E'; $iFg = HC '#6B7280'
    }
    if ($tab -eq "NTFS") {
        $tabNTFS.BackColor = $aBg; $tabNTFS.ForeColor = $aFg; $tabNTFS.FlatAppearance.BorderSize = 0
        $tabSP.BackColor   = $iBg; $tabSP.ForeColor   = $iFg
        $pnlIn.Visible = $true;  $pnlSP.Visible = $false
    } else {
        $tabSP.BackColor   = $aBg; $tabSP.ForeColor   = $aFg
        $tabNTFS.BackColor = $iBg; $tabNTFS.ForeColor = $iFg; $tabNTFS.FlatAppearance.BorderSize = 0
        $pnlIn.Visible = $false; $pnlSP.Visible = $true
    }
    Do-Layout
}

$tabNTFS.add_Click({ Set-ActiveTab "NTFS" })
$tabSP.add_Click({   Set-ActiveTab "SP"   })

# ── Eingabefelder ────────────────────────────────────────────
function New-Label($text, $bold) {
    $l = [System.Windows.Forms.Label]::new()
    $l.Text      = $text
    $l.Font      = if ($bold) { $FB } else { $FS }
    $l.ForeColor = $C_TXT
    $l.BackColor = [System.Drawing.Color]::Transparent
    $l.AutoSize  = $true
    return $l
}

function New-TB($placeholder) {
    $t = [System.Windows.Forms.TextBox]::new()
    $t.Height      = $BTN_H
    $t.BackColor   = $C_INPUT
    $t.BorderStyle = "FixedSingle"
    $t.Font        = $FN
    if ($placeholder) {
        $t.Text      = $placeholder
        $t.ForeColor = $C_LOW
    } else {
        $t.ForeColor = $C_TXT
    }
    return $t
}

function New-Btn($text) {
    $b = [System.Windows.Forms.Button]::new()
    $b.Text      = $text
    $b.Size      = [System.Drawing.Size]::new($BTN_W, $BTN_H)
    $b.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $b.Font      = $FN
    $b.BackColor = $C_BTN
    $b.ForeColor = $C_TXT
    $b.Cursor    = "Hand"
    $b.FlatAppearance.BorderColor = $C_BORD2
    $b.FlatAppearance.BorderSize  = 1
    return $b
}

# Zeile 1: Start Path
$lblPath   = New-Label $script:STR.LblPath $true
$tbPath    = New-TB $null
$btnBrowse = New-Btn $script:STR.LblBrowse
$pnlIn.Controls.Add($lblPath)
$pnlIn.Controls.Add($tbPath)
$pnlIn.Controls.Add($btnBrowse)

# Zeile 2: Output File
$lblOut  = New-Label $script:STR.LblOutput $true
$tbOut   = New-TB "(automatic — saved to script folder)"
$btnSave = New-Btn $script:STR.LblSaveAs
$pnlIn.Controls.Add($lblOut)
$pnlIn.Controls.Add($tbOut)
$pnlIn.Controls.Add($btnSave)

# Zeile 3: Depth
$lblDepth = New-Label $script:STR.LblDepth $true
$pnlIn.Controls.Add($lblDepth)

$numDepth = [System.Windows.Forms.NumericUpDown]::new()
$numDepth.Size      = [System.Drawing.Size]::new(68, $BTN_H)
$numDepth.Minimum   = 0
$numDepth.Maximum   = 100
$numDepth.Value     = 0
$numDepth.BackColor = $C_INPUT
$numDepth.ForeColor = $C_TXT
$numDepth.Font      = $FN
$pnlIn.Controls.Add($numDepth)

$lblUnlim = New-Label "  0 = unlimited" $false
$lblUnlim.ForeColor = $C_LOW
$pnlIn.Controls.Add($lblUnlim)

# Checkbox
$chk = [System.Windows.Forms.CheckBox]::new()
$chk.Text      = $script:STR.ChkBrowser
$chk.Font      = $FN
$chk.ForeColor = $C_MID
$chk.BackColor = [System.Drawing.Color]::Transparent
$chk.Checked   = $true
$chk.AutoSize  = $true
$pnlIn.Controls.Add($chk)

# Separator (zwischen Feldern und Progress)
$sepMid = [System.Windows.Forms.Panel]::new()
$sepMid.BackColor = $C_BORDER
$pnlIn.Controls.Add($sepMid)

# ── Progress ─────────────────────────────────────────────────
$lblStatus = [System.Windows.Forms.Label]::new()
$lblStatus.Text      = $script:STR.LblReady
$lblStatus.Font      = $FN
$lblStatus.ForeColor = $C_MID
$lblStatus.BackColor = [System.Drawing.Color]::Transparent
$pnlIn.Controls.Add($lblStatus)

$progBar = [System.Windows.Forms.ProgressBar]::new()
$progBar.Style   = "Continuous"
$progBar.Minimum = 0
$progBar.Maximum = 100
$pnlIn.Controls.Add($progBar)

$lblStats = [System.Windows.Forms.Label]::new()
$lblStats.Text      = ""
$lblStats.Font      = $FS
$lblStats.ForeColor = $C_MID
$lblStats.BackColor = [System.Drawing.Color]::Transparent
$pnlIn.Controls.Add($lblStats)

# ── Layout-Funktion (absolut, sauber) ────────────────────────
function Do-Layout {
    $cw = $form.ClientSize.Width
    $ch = $form.ClientSize.Height

    $x      = $PAD
    $tbW    = $cw - $PAD * 2 - $BTN_W - 8   # Textbox-Breite
    $fullW  = $cw - $PAD * 2                 # Volle Inhaltsbreite

    # Header & Tab-Bar
    $pnlHdr.SetBounds(0, 0, $cw, $HDR_H)
    if ($btnTheme) { $btnTheme.Location = [System.Drawing.Point]::new($cw - $btnTheme.Width - 14, [int](($HDR_H - $btnTheme.Height) / 2)) }
    if ($btnLang)  { $btnLang.Location  = [System.Drawing.Point]::new($cw - $btnTheme.Width - $btnLang.Width - 22, [int](($HDR_H - $btnLang.Height) / 2)) }
    $sepHdr.SetBounds(0, $HDR_H, $cw, 1)
    $pnlTabs.SetBounds(0, $HDR_H + 1, $cw, 36)
    $tabNTFS.Location  = [System.Drawing.Point]::new(0, 0)
    $tabSP.Location    = [System.Drawing.Point]::new(120, 0)
    $tabLineW = if ($script:activeTab -eq "NTFS") { 120 } else { 140 }
    $tabLineX = if ($script:activeTab -eq "NTFS") { 0 } else { 120 }
    $tabLine.SetBounds($tabLineX, 34, $tabLineW, 2)
    $tabSep.SetBounds(0, $HDR_H + 37, $cw, 1)
    $contentY = $HDR_H + 38
    $contentH = $ch - $FTR_H - $contentY - 2
    $pnlIn.SetBounds(0, $contentY, $cw, $contentH)
    $pnlSP.SetBounds(0, $contentY, $cw, $contentH)
    # SP controls layout
    $spW = $cw - $PAD * 2
    $spLblConn.Location       = [System.Drawing.Point]::new($PAD, 14)
    $spLblConnDetail.SetBounds($PAD, 36, $spW, 20)
    $spLblUrl.Location        = [System.Drawing.Point]::new($PAD, 66)
    $spTbUrl.SetBounds($PAD, 88, $spW - 170, 26)
    $btnSPSetup.Location      = [System.Drawing.Point]::new($spW - 160 + $PAD, 88)
    $spLblDepth.Location      = [System.Drawing.Point]::new($PAD, 126)
    $spNumDepth.Location      = [System.Drawing.Point]::new($PAD, 148)
    $spLblUnlim.Location      = [System.Drawing.Point]::new($PAD + 74, 152)
    $spChk.Location           = [System.Drawing.Point]::new($PAD, 186)

    # Footer & Trennlinie (von unten)
    $sepFtr.SetBounds(0, $ch - $FTR_H - 1, $cw, 1)
    $pnlFtr.SetBounds(0, $ch - $FTR_H, $cw, $FTR_H)
    # Auto-layout footer buttons left to right
    $fGap = 8
    $fX   = 16
    $fBtnY = [int](($FTR_H - 36) / 2)
    $fBtns = @($btnStart, $btnCancel, $btnOpenLast, $btnChangelog, $btnCompare)
    $fTotalW = ($fBtns | Measure-Object -Property Width -Sum).Sum + $fGap * ($fBtns.Count - 1)
    $fX = [int](($cw - $fTotalW) / 2)
    foreach ($fb in $fBtns) {
        $fb.Location = [System.Drawing.Point]::new($fX, $fBtnY)
        $fX += $fb.Width + $fGap
    }


    # Inhaltsbereich Y-Start — NTFS panel is now self-contained, Y starts at 0
    $contentTop = 0

    # Zeile 1: Start Path  (Y = contentTop + 16)
    $y1 = $contentTop + 16
    $lblPath.Location  = [System.Drawing.Point]::new($x, $y1)
    $tbPath.SetBounds($x, ($y1 + 22), $tbW, $BTN_H)
    $btnBrowse.Location = [System.Drawing.Point]::new($x + $tbW + 8, $y1 + 22)

    # Zeile 2: Output File (Y = y1 + ROW_H)
    $y2 = $y1 + $ROW_H
    $lblOut.Location  = [System.Drawing.Point]::new($x, $y2)
    $tbOut.SetBounds($x, ($y2 + 22), $tbW, $BTN_H)
    $btnSave.Location = [System.Drawing.Point]::new($x + $tbW + 8, $y2 + 22)

    # Zeile 3: Max Depth (Y = y2 + ROW_H)
    $y3 = $y2 + $ROW_H
    $lblDepth.Location  = [System.Drawing.Point]::new($x, $y3)
    $numDepth.Location  = [System.Drawing.Point]::new($x, $y3 + 22)
    $lblUnlim.Location  = [System.Drawing.Point]::new($x + 74, $y3 + 26)

    # Checkbox (Y = y3 + ROW_H - 10)
    $y4 = $y3 + $ROW_H - 4
    $chk.Location = [System.Drawing.Point]::new($x, $y4)

    # Separator Mitte
    $sepMidY = $y4 + 32
    $sepMid.SetBounds(0, $sepMidY, $cw, 1)

    # Progress-Bereich
    $progY = $sepMidY + 14
    $lblStatus.SetBounds($x, $progY, $fullW, 20)
    $progBar.SetBounds($x, ($progY + 24), $fullW, 14)
    $lblStats.SetBounds($x, ($progY + 42), $fullW, 18)
}

$form.add_Load({
    Do-Layout
    # Pin all header items to same Y = 14 (visual baseline for 54px header)
    $lblTitle.Location   = [System.Drawing.Point]::new(18, 13)
    # Bottom-align smaller labels to title baseline
    $botY = $lblTitle.Bottom - $lblVersion.Height - 2
    $lblVersion.Location = [System.Drawing.Point]::new($lblTitle.Right + 8, $botY)
    $lblSub.Location     = [System.Drawing.Point]::new($lblVersion.Right + 10, $lblTitle.Bottom - $lblSub.Height - 2)
})
$form.add_Resize({ Do-Layout })

# ── Events ───────────────────────────────────────────────────
$btnBrowse.add_Click({
    $fbd = [System.Windows.Forms.FolderBrowserDialog]::new()
    $fbd.Description = "Select folder for NTFS analysis"
    if ($tbPath.Text -and (Test-Path $tbPath.Text)) { $fbd.SelectedPath = $tbPath.Text }
    if ($fbd.ShowDialog() -eq "OK") {
        $tbPath.Text      = $fbd.SelectedPath
        $tbPath.ForeColor = $C_TXT
    }
})

$btnSave.add_Click({
    $sfd = [System.Windows.Forms.SaveFileDialog]::new()
    $sfd.Filter   = "HTML File (*.html)|*.html"
    $sfd.FileName = "ACLens_Report_$(Get-Date -Format 'yyyy-MM-dd').html"
    if ($tbPath.Text -and (Test-Path $tbPath.Text)) { $sfd.InitialDirectory = $tbPath.Text }
    if ($sfd.ShowDialog() -eq "OK") {
        $tbOut.Text      = $sfd.FileName
        $tbOut.ForeColor = $C_TXT
    }
})

$script:cancelFlag = $false
    $script:lastReport  = ""
$btnCancel.add_Click({ if ($script:scanning) { $script:cancelFlag = $true } })

$btnStart.add_Click({
    # Route to correct scanner based on active tab
    if ($script:activeTab -eq "SP") {
        $siteUrl = $spTbUrl.Text.Trim()
        if ([string]::IsNullOrEmpty($siteUrl) -or $siteUrl -like "*yourtenant*") {
            [System.Windows.Forms.MessageBox]::Show("Please enter a valid SharePoint Site URL.", "Invalid URL", "OK", "Warning") | Out-Null
            return
        }
        $creds = Get-SPCredentials
        if (-not $creds.IsConfigured) {
            [System.Windows.Forms.MessageBox]::Show("SharePoint is not configured yet.`nPlease use 'Setup / Reconnect' first.", "Not configured", "OK", "Warning") | Out-Null
            return
        }
        $ts      = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
        $outFile = Join-Path $PSScriptRoot "ACLens_SP_Report_$ts.html"
        $btnStart.Enabled  = $false
        $progBar.Value     = 0
        $lblStats.Text     = ""
        try {
            Start-SPScan -SiteUrl $siteUrl -MaxDepth ([int]$spNumDepth.Value) `
                -OutputPath $outFile -StatusLabel $lblStatus -ProgressBar $progBar `
                -StatsLabel $lblStats -FormRef $form
            $script:lastReport = $outFile
            $btnOpenLast.Enabled = $true
            if ($spChk.Checked) { Start-Process $outFile }
        } catch {
            $lblStatus.Text      = "Error: $($_.Exception.Message)"
            $lblStatus.ForeColor = $C_ERR
        } finally {
            $btnStart.Enabled  = $true
        }
        return
    }

    $rootPath = $tbPath.Text.Trim()
    if ([string]::IsNullOrEmpty($rootPath) -or -not (Test-Path $rootPath -PathType Container)) {
        [System.Windows.Forms.MessageBox]::Show(
            $script:STR.MsgInvalidPath, $script:STR.MsgInvalidPathT, "OK", "Warning") | Out-Null
        return
    }

    $outFile = $tbOut.Text.Trim()
    if ([string]::IsNullOrEmpty($outFile) -or $outFile -like "*(automatic*") {
        $ts      = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
        $outFile = Join-Path $PSScriptRoot "ACLens_Report_$ts.html"
    }
    if (-not [System.IO.Path]::IsPathRooted($outFile)) {
        $outFile = Join-Path $PSScriptRoot $outFile
    }

    $maxDepth          = [int]$numDepth.Value
    $btnStart.Enabled  = $false
    $script:cancelFlag = $false
    $progBar.Value     = 0
    $lblStats.Text     = ""
    $lblStatus.ForeColor = $C_MID

    try {
        $lblStatus.Text = $script:STR.StepScanFolders
        $form.Refresh()

        $allFolders = Get-AllFolders -RootPath $rootPath -MaxDepth $maxDepth
        $total      = $allFolders.Count

        $lblStatus.Text = "Step 2/3: Reading NTFS permissions ($total folders)..."
        $form.Refresh()

        $folderData    = [System.Collections.Generic.List[object]]::new()
        $prevAcl       = $null
        $rootPathClean = $rootPath.TrimEnd('\')
        $counter       = 0

        foreach ($folderPath in $allFolders) {
            if ($script:cancelFlag) {
                $lblStatus.Text      = $script:STR.StepCancelled
                $lblStatus.ForeColor = $C_WARN
                $progBar.Value       = 0
                $btnStart.Enabled    = $true
                return
            }

            $counter++
            $pct = [int](($counter / $total) * 100)
            if ($progBar.Value -ne $pct) { $progBar.Value = $pct }
            if ($counter % 5 -eq 0 -or $counter -eq $total) {
                $short = if ($folderPath.Length -gt 90) {
                    "..." + $folderPath.Substring($folderPath.Length - 87) } else { $folderPath }
                $lblStatus.Text = "Step 2/3: $counter / $total  —  $short"
                $form.Refresh()
            }

            $relativePath = $folderPath.Substring($rootPathClean.Length).TrimStart('\')
            $depth        = if ($relativePath -eq "") { 0 } else { ($relativePath -split '\\').Count }
            $aclData      = Get-FolderACL -FolderPath $folderPath
            $allInherited = (@($aclData.Rules | Where-Object { -not $_.IsInherited }).Count -eq 0)

            $hasChanges = $true
            if ($depth -gt 0 -and ($null -ne $prevAcl) -and
                ($null -eq $aclData.Error) -and ($null -eq $prevAcl.Error)) {
                $sameRules   = Compare-ACLRules -Rules1 $aclData.Rules -Rules2 $prevAcl.Rules
                $sameInherit = ($aclData.InheritanceEnabled -eq $prevAcl.InheritanceEnabled)
                $hasChanges  = -not ($sameRules -and $sameInherit)
            } elseif ($depth -eq 0) { $hasChanges = $true }

            $folderData.Add([PSCustomObject]@{
                Path                    = $folderPath
                Depth                   = $depth
                Owner                   = $aclData.Owner
                InheritanceEnabled      = $aclData.InheritanceEnabled
                AreAccessRulesProtected = $aclData.AreAccessRulesProtected
                Rules                   = $aclData.Rules
                Error                   = $aclData.Error
                HasChanges              = $hasChanges
                AllInherited            = $allInherited
            })
            $prevAcl = $aclData
        }

        $lblStatus.Text = $script:STR.StepGenReport
        $progBar.Value  = 99
        $form.Refresh()

        New-HTMLReport -FolderData $folderData.ToArray() -RootPath $rootPath -OutputPath $outFile
        $progBar.Value = 100

        $changedCount   = ($folderData | Where-Object { $_.HasChanges   -eq $true }).Count
        $errorCount     = ($folderData | Where-Object { $null -ne $_.Error }).Count
        $inheritedCount = ($folderData | Where-Object {
            $_.AllInherited -eq $true -and $_.HasChanges -eq $false }).Count

        $script:lastReport       = $outFile
        $btnOpenLast.Enabled     = $true

        # Save JSON snapshot for future comparison
        $jsonFile = [System.IO.Path]::ChangeExtension($outFile, '.json')
        $snapshot = $folderData.ToArray() | Select-Object Path, Depth, Owner, InheritanceEnabled, AreAccessRulesProtected, HasChanges, AllInherited, Error,
            @{N='Rules';E={ $_.Rules | Select-Object Identity, AccessType, Rights, IsInherited, InheritanceFlags, PropagationFlags }}
        $snapshot | ConvertTo-Json -Depth 6 | Out-File -FilePath $jsonFile -Encoding UTF8
        $script:lastJson = $jsonFile

        $lblStatus.Text          = "Done!   Report saved to: $outFile"
        $lblStatus.ForeColor = $C_OK
        $lblStats.Text       = "Total: $total   |   Changed: $changedCount   |   Inherited only: $inheritedCount   |   Errors: $errorCount"

        if ($chk.Checked) { Start-Process $outFile }

    } catch {
        $lblStatus.Text      = "Error: $($_.Exception.Message)"
        $lblStatus.ForeColor = $C_ERR
    } finally {
        $btnStart.Enabled  = $true
    }
})


# ── SharePoint Tab Controls ───────────────────────────────────
$spCredsFile = Join-Path $PSScriptRoot "sp_credentials.xml"

function Get-SPCredentials {
    if (Test-Path $spCredsFile) {
        try {
            $xml = [xml](Get-Content $spCredsFile -Raw)
            return [PSCustomObject]@{
                TenantId     = $xml.ACLens.TenantId
                ClientId     = $xml.ACLens.ClientId
                IsConfigured = ($xml.ACLens.TenantId -ne "")
            }
        } catch { }
    }
    return [PSCustomObject]@{ TenantId = ""; ClientId = ""; IsConfigured = $false }
}

function Save-SPCredentials($tenantId, $clientId, $clientSecret) {
    $secureSecret    = ConvertTo-SecureString $clientSecret -AsPlainText -Force
    $encryptedSecret = $secureSecret | ConvertFrom-SecureString
    @"
<?xml version="1.0" encoding="utf-8"?>
<ACLens>
  <TenantId>$tenantId</TenantId>
  <ClientId>$clientId</ClientId>
  <ClientSecret>$encryptedSecret</ClientSecret>
</ACLens>
"@ | Out-File -FilePath $spCredsFile -Encoding UTF8 -Force
}

# SP: Connection status
$spLblConn = [System.Windows.Forms.Label]::new()
$spLblConn.Font = $FB; $spLblConn.BackColor = [System.Drawing.Color]::Transparent; $spLblConn.AutoSize = $true
$pnlSP.Controls.Add($spLblConn)

$spLblConnDetail = [System.Windows.Forms.Label]::new()
$spLblConnDetail.Font = $FS; $spLblConnDetail.ForeColor = HC '#9CA3AF'
$spLblConnDetail.BackColor = [System.Drawing.Color]::Transparent
$pnlSP.Controls.Add($spLblConnDetail)

# SP: Site URL
$spLblUrl = [System.Windows.Forms.Label]::new()
$spLblUrl.Text = $script:STR.SpLblUrl; $spLblUrl.Font = $FB
$spLblUrl.ForeColor = HC '#E2E8F0'; $spLblUrl.BackColor = [System.Drawing.Color]::Transparent; $spLblUrl.AutoSize = $true
$pnlSP.Controls.Add($spLblUrl)

$spTbUrl = [System.Windows.Forms.TextBox]::new()
$spTbUrl.BackColor = HC '#2D2D3F'; $spTbUrl.ForeColor = HC '#6B7280'
$spTbUrl.Text = "https://yourtenant.sharepoint.com/sites/yoursite"
$spTbUrl.BorderStyle = "FixedSingle"; $spTbUrl.Font = $FN; $spTbUrl.Height = 26
$pnlSP.Controls.Add($spTbUrl)

# SP: Depth
$spLblDepth = [System.Windows.Forms.Label]::new()
$spLblDepth.Text = $script:STR.SpLblDepth; $spLblDepth.Font = $FB
$spLblDepth.ForeColor = HC '#E2E8F0'; $spLblDepth.BackColor = [System.Drawing.Color]::Transparent; $spLblDepth.AutoSize = $true
$pnlSP.Controls.Add($spLblDepth)

$spNumDepth = [System.Windows.Forms.NumericUpDown]::new()
$spNumDepth.Size = [System.Drawing.Size]::new(68, 26); $spNumDepth.Minimum = 0; $spNumDepth.Maximum = 100; $spNumDepth.Value = 0
$spNumDepth.BackColor = HC '#2D2D3F'; $spNumDepth.ForeColor = HC '#E2E8F0'; $spNumDepth.Font = $FN
$pnlSP.Controls.Add($spNumDepth)

$spLblUnlim = [System.Windows.Forms.Label]::new()
$spLblUnlim.Text = $script:STR.LblUnlimited; $spLblUnlim.Font = $FS
$spLblUnlim.ForeColor = HC '#6B7280'; $spLblUnlim.BackColor = [System.Drawing.Color]::Transparent; $spLblUnlim.AutoSize = $true
$pnlSP.Controls.Add($spLblUnlim)

# SP: Checkbox
$spChk = [System.Windows.Forms.CheckBox]::new()
$spChk.Text = $script:STR.SpChkBrowser; $spChk.Font = $FN
$spChk.ForeColor = HC '#9CA3AF'; $spChk.BackColor = [System.Drawing.Color]::Transparent
$spChk.Checked = $true; $spChk.AutoSize = $true
$pnlSP.Controls.Add($spChk)

# SP: Setup button
$btnSPSetup = [System.Windows.Forms.Button]::new()
$btnSPSetup.Text = $script:STR.SpBtnSetup; $btnSPSetup.Size = [System.Drawing.Size]::new(160, 26)
$btnSPSetup.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $btnSPSetup.Font = $FN
$btnSPSetup.BackColor = HC '#7C3AED'; $btnSPSetup.ForeColor = [System.Drawing.Color]::White; $btnSPSetup.Cursor = "Hand"
$btnSPSetup.FlatAppearance.BorderColor = HC '#4C1D95'; $btnSPSetup.FlatAppearance.BorderSize = 1
    $btnSPSetup.UseVisualStyleBackColor = $false
$pnlSP.Controls.Add($btnSPSetup)

function Update-SPStatus {
    $creds = Get-SPCredentials
    if ($creds.IsConfigured) {
        $spLblConn.Text      = $script:STR.SpConnected
        $spLblConn.ForeColor = HC '#34D399'
        $spLblConnDetail.Text = "Tenant: $($creds.TenantId)  |  Client: $($creds.ClientId)"
    } else {
        $spLblConn.Text      = $script:STR.SpNotConfigured
        $spLblConn.ForeColor = HC '#F87171'
        $spLblConnDetail.Text = $script:STR.SpConnectHint
    }
}

Update-SPStatus
$btnSPSetup.add_Click({ Show-SPSetupWizard })

# ── SharePoint Scanner Functions ─────────────────────────────
function Get-SPAccessToken {
    $credsFile = Join-Path $PSScriptRoot "sp_credentials.xml"
    if (-not (Test-Path $credsFile)) { throw "SharePoint not configured. Use Setup / Reconnect first." }
    $xml = [xml](Get-Content $credsFile -Raw)
    $tenantId  = $xml.ACLens.TenantId
    $clientId  = $xml.ACLens.ClientId
    $secSecret = $xml.ACLens.ClientSecret | ConvertTo-SecureString
    $bstr      = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secSecret)
    $secret    = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
    [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)

    $body = @{
        grant_type    = "client_credentials"
        client_id     = $clientId
        client_secret = $secret
        scope         = "https://graph.microsoft.com/.default"
    }
    $resp = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" `
        -Method Post -Body $body -ContentType "application/x-www-form-urlencoded"
    return $resp.access_token
}

function Invoke-GraphAPI {
    param([string]$Uri, [string]$Token, [string]$Method = "GET", [string]$Body = $null)
    $headers = @{ Authorization = "Bearer $Token"; "Content-Type" = "application/json" }
    $params  = @{ Uri = $Uri; Headers = $headers; Method = $Method; ErrorAction = "Stop" }
    if ($Body) { $params.Body = $Body }
    $resp    = Invoke-RestMethod @params
    return $resp
}

function Get-GraphPagedResults {
    param([string]$Uri, [string]$Token)
    $results = [System.Collections.Generic.List[object]]::new()
    $nextUri = $Uri
    while ($nextUri) {
        $page    = Invoke-GraphAPI -Uri $nextUri -Token $Token
        if ($page.value) { foreach ($item in $page.value) { $results.Add($item) } }
        $nextUri = $page.'@odata.nextLink'
    }
    return $results.ToArray()
}

function Get-SPSitePermissions {
    param([string]$SiteId, [string]$Token)
    $perms = @()
    try {
        $resp = Invoke-GraphAPI -Uri "https://graph.microsoft.com/v1.0/sites/$SiteId/permissions" -Token $Token
        foreach ($p in $resp.value) {
            $roles     = $p.roles -join ", "
            $grantedTo = if ($p.grantedToIdentitiesV2) {
                ($p.grantedToIdentitiesV2 | ForEach-Object {
                    if ($_.user)  { $_.user.displayName }
                    elseif ($_.group) { $_.group.displayName }
                    else { "Unknown" }
                }) -join "; "
            } elseif ($p.grantedToV2) {
                if ($p.grantedToV2.user)  { $p.grantedToV2.user.displayName }
                elseif ($p.grantedToV2.group) { $p.grantedToV2.group.displayName }
                else { "Unknown" }
            } else { "Unknown" }

            $perms += [PSCustomObject]@{
                GrantedTo   = $grantedTo
                Roles       = $roles
                Type        = if ($p.grantedToV2.user) { "User" } elseif ($p.grantedToV2.group) { "Group" } else { "App" }
                InviteToken = $p.id
            }
        }
    } catch { }
    return $perms
}

function Get-SPListPermissions {
    param([string]$SiteId, [string]$ListId, [string]$Token)
    # SharePoint REST for list-level permissions (Graph doesn't expose these directly)
    $perms = @()
    try {
        $siteResp = Invoke-GraphAPI -Uri "https://graph.microsoft.com/v1.0/sites/$SiteId" -Token $Token
        $hostname  = ([System.Uri]$siteResp.webUrl).Host
        $sitePath  = ([System.Uri]$siteResp.webUrl).AbsolutePath
        $restUrl   = "https://$hostname/_api/web/lists(guid'$ListId')/roleassignments?`$expand=Member,RoleDefinitionBindings"
        $headers   = @{ Authorization = "Bearer $Token"; Accept = "application/json;odata=nometadata" }
        $resp      = Invoke-RestMethod -Uri $restUrl -Headers $headers -ErrorAction Stop
        foreach ($ra in $resp.value) {
            $memberName = $ra.Member.Title
            $roles      = ($ra.RoleDefinitionBindings | ForEach-Object { $_.Name }) -join ", "
            $memberType = switch ($ra.Member.PrincipalType) {
                1  { "User" }
                 4 { "SecurityGroup" }
                8  { "SharePoint Group" }
                default { "Unknown" }
            }
            $perms += [PSCustomObject]@{
                GrantedTo       = $memberName
                Roles           = $roles
                Type            = $memberType
                UniquePerms     = $true
            }
        }
    } catch { }
    return $perms
}

function Get-SPGroupMembers {
    param([string]$SiteId, [string]$Token)
    $groups = [System.Collections.Generic.List[object]]::new()
    try {
        $siteResp = Invoke-GraphAPI -Uri "https://graph.microsoft.com/v1.0/sites/$SiteId" -Token $Token
        $hostname  = ([System.Uri]$siteResp.webUrl).Host
        $restUrl   = "https://$hostname/_api/web/sitegroups?`$expand=Users"
        $headers   = @{ Authorization = "Bearer $Token"; Accept = "application/json;odata=nometadata" }
        $resp      = Invoke-RestMethod -Uri $restUrl -Headers $headers -ErrorAction Stop
        foreach ($grp in $resp.value) {
            $members = ($grp.Users | ForEach-Object { $_.Title }) -join "; "
            $groups.Add([PSCustomObject]@{
                Name    = $grp.Title
                Members = $members
                Count   = $grp.Users.Count
            })
        }
    } catch { }
    return $groups.ToArray()
}

function Get-SPExternalSharing {
    param([string]$SiteId, [string]$Token)
    try {
        $site = Invoke-GraphAPI -Uri "https://graph.microsoft.com/v1.0/sites/$SiteId" -Token $Token
        $policy = $site.sharingCapability
        $result = switch ($policy) {
            "disabled"                        { "Disabled" }
            "externalUserSharingOnly"         { "External users only" }
            "externalUserAndGuestSharing"     { "External users and guests" }
            "existingExternalUserSharingOnly" { "Existing external users only" }
            default                           { if ($policy) { $policy } else { "Unknown" } }
        }
        return $result
    } catch { return "Unknown" }
}

function Get-SPFolderTree {
    param([string]$SiteId, [string]$DriveId, [string]$ItemId = "root", [int]$MaxDepth, [int]$CurrentDepth = 0, [string]$Token)
    $results = [System.Collections.Generic.List[object]]::new()
    if ($MaxDepth -gt 0 -and $CurrentDepth -ge $MaxDepth) { return $results.ToArray() }
    try {
        $uri = if ($ItemId -eq "root") {
            "https://graph.microsoft.com/v1.0/sites/$SiteId/drives/$DriveId/root/children"
        } else {
            "https://graph.microsoft.com/v1.0/sites/$SiteId/drives/$DriveId/items/$ItemId/children"
        }
        $children = Get-GraphPagedResults -Uri $uri -Token $Token
        foreach ($child in $children) {
            if ($child.folder) {
                # Get permissions for this folder
                $permUri  = "https://graph.microsoft.com/v1.0/sites/$SiteId/drives/$DriveId/items/$($child.id)/permissions"
                $permResp = Invoke-GraphAPI -Uri $permUri -Token $Token -ErrorAction SilentlyContinue
                $perms    = @()
                if ($permResp -and $permResp.value) {
                    foreach ($p in $permResp.value) {
                        $grantedTo = "Unknown"
                        if ($p.grantedToV2) {
                            if ($p.grantedToV2.user)        { $grantedTo = $p.grantedToV2.user.displayName }
                            elseif ($p.grantedToV2.group)   { $grantedTo = $p.grantedToV2.group.displayName }
                            elseif ($p.grantedToV2.siteUser){ $grantedTo = $p.grantedToV2.siteUser.displayName }
                        }
                        $roles = if ($p.roles) { $p.roles -join ", " } else { "Custom" }
                        $perms += [PSCustomObject]@{
                            GrantedTo   = $grantedTo
                            Roles       = $roles
                            IsInherited = ($p.inheritedFrom -ne $null)
                            Link        = ($p.link -ne $null)
                        }
                    }
                }
                $hasUniquePerms = ($perms | Where-Object { -not $_.IsInherited }).Count -gt 0
                $results.Add([PSCustomObject]@{
                    Path            = $child.name
                    FullPath        = $child.parentReference.path + "/" + $child.name
                    WebUrl          = $child.webUrl
                    Depth           = $CurrentDepth
                    HasUniquePerms  = $hasUniquePerms
                    Permissions     = $perms
                    Error           = $null
                })
                # Recurse
                $sub = Get-SPFolderTree -SiteId $SiteId -DriveId $DriveId -ItemId $child.id `
                    -MaxDepth $MaxDepth -CurrentDepth ($CurrentDepth + 1) -Token $Token
                foreach ($s in $sub) { $results.Add($s) }
            }
        }
    } catch {
        $results.Add([PSCustomObject]@{
            Path = "Error"; FullPath = ""; WebUrl = ""; Depth = $CurrentDepth
            HasUniquePerms = $false; Permissions = @(); Error = $_.Exception.Message
        })
    }
    return $results.ToArray()
}

function Start-SPScan {
    param(
        [string]$SiteUrl,
        [int]$MaxDepth,
        [string]$OutputPath,
        [System.Windows.Forms.Label]$StatusLabel,
        [System.Windows.Forms.ProgressBar]$ProgressBar,
        [System.Windows.Forms.Label]$StatsLabel,
        [System.Windows.Forms.Form]$FormRef
    )

    $StatusLabel.Text      = "Authenticating..."
    $StatusLabel.ForeColor = HC '#9CA3AF'
    $FormRef.Refresh()

    $token = Get-SPAccessToken

    # Resolve site ID from URL
    $StatusLabel.Text = "Resolving site..."
    $FormRef.Refresh()

    $encodedUrl  = [System.Uri]::EscapeDataString($SiteUrl)
    $siteInfo    = Invoke-GraphAPI -Uri "https://graph.microsoft.com/v1.0/sites?`$filter=webUrl eq '$encodedUrl'" -Token $token
    if (-not $siteInfo.value -or $siteInfo.value.Count -eq 0) {
        # Try direct lookup
        $uri2    = $SiteUrl -replace "https://", "" -replace "\.sharepoint\.com", ".sharepoint.com:"
        $siteInfo2 = Invoke-GraphAPI -Uri "https://graph.microsoft.com/v1.0/sites/$uri2" -Token $token
        $site    = $siteInfo2
    } else {
        $site = $siteInfo.value[0]
    }
    $siteId = $site.id

    # Get site-level permissions
    $StatusLabel.Text = "Reading site permissions..."
    $FormRef.Refresh()
    $sitePerms    = Get-SPSitePermissions -SiteId $siteId -Token $token
    $siteGroups   = Get-SPGroupMembers -SiteId $siteId -Token $token
    $extSharing   = Get-SPExternalSharing -SiteId $siteId -Token $token

    # Get drives (document libraries)
    $StatusLabel.Text = "Scanning document libraries..."
    $FormRef.Refresh()
    $drives = Get-GraphPagedResults -Uri "https://graph.microsoft.com/v1.0/sites/$siteId/drives" -Token $token

    $allItems  = [System.Collections.Generic.List[object]]::new()
    $driveData = [System.Collections.Generic.List[object]]::new()
    $totalDrives = $drives.Count
    $driveIdx    = 0

    foreach ($drive in $drives) {
        $driveIdx++
        $pct = [int](($driveIdx / $totalDrives) * 90)
        $ProgressBar.Value = $pct
        $StatusLabel.Text  = "Scanning library $driveIdx / $totalDrives : $($drive.name)"
        $FormRef.Refresh()

        # Get library-level permissions
        $listId    = $drive.list.id
        $listPerms = if ($listId) { Get-SPListPermissions -SiteId $siteId -ListId $listId -Token $token } else { @() }

        # Get folder tree
        $folders = Get-SPFolderTree -SiteId $siteId -DriveId $drive.id `
            -MaxDepth $MaxDepth -CurrentDepth 0 -Token $token

        $driveData.Add([PSCustomObject]@{
            Name        = $drive.name
            DriveType   = $drive.driveType
            WebUrl      = $drive.webUrl
            Permissions = $listPerms
            Folders     = $folders
        })
        foreach ($f in $folders) { $allItems.Add($f) }
    }

    $ProgressBar.Value = 99
    $StatusLabel.Text  = "Generating HTML report..."
    $FormRef.Refresh()

    $result = New-SPHTMLReport -SiteUrl $SiteUrl -SiteId $siteId -SiteName $site.displayName `
        -SitePerms $sitePerms -Groups $siteGroups -ExternalSharing $extSharing `
        -Drives $driveData -OutputPath $OutputPath

    $ProgressBar.Value   = 100
    $StatusLabel.Text    = "Done!  Report saved to: $OutputPath"
    $StatusLabel.ForeColor = HC '#34D399'
    $StatsLabel.Text     = "Libraries: $($drives.Count)   |   Folders scanned: $($allItems.Count)   |   External Sharing: $extSharing"

    # Save SP JSON snapshot for comparison
    $jsonOut  = [System.IO.Path]::ChangeExtension($OutputPath, '.json')
    $snapshot = $allItems.ToArray() | Select-Object FullPath, WebUrl, Depth, HasUniquePerms, Error,
        @{N='Permissions';E={ $_.Permissions | Select-Object GrantedTo, Roles, IsInherited, Link }}
    $snapshot | ConvertTo-Json -Depth 5 | Out-File -FilePath $jsonOut -Encoding UTF8

    return $result
}


function Show-SPSetupWizard {
    $wiz = [System.Windows.Forms.Form]::new()
    $wiz.Text            = $script:STR.WizTitle
    $wiz.ClientSize      = [System.Drawing.Size]::new(660, 480)
    $wiz.StartPosition   = "CenterParent"
    $wiz.FormBorderStyle = "FixedDialog"
    $wiz.MaximizeBox     = $false
    $wizBg  = if ($script:lightMode) { HC '#F5F5F5' } else { HC '#1E1E2E' }
    $wizHdr = if ($script:lightMode) { HC '#ECECEC' } else { HC '#1A1A2E' }
    $wizSep = if ($script:lightMode) { HC '#CCCCCC' } else { HC '#4B5563' }
    $wizPnl = if ($script:lightMode) { HC '#F5F5F5' } else { HC '#1E1E2E' }
    $wizTxt = if ($script:lightMode) { HC '#111111' } else { HC '#E2E8F0' }
    $wizMid = if ($script:lightMode) { HC '#555555' } else { HC '#9CA3AF' }
    $wizLow = if ($script:lightMode) { HC '#6B7280' } else { HC '#6B7280' }
    $wiz.BackColor = $wizBg
    $wiz.Font      = $FN

    # Header
    $wHdr = [System.Windows.Forms.Panel]::new()
    $wHdr.SetBounds(0, 0, 660, 64); $wHdr.BackColor = $wizHdr
    $wiz.Controls.Add($wHdr)

    $wTitle = [System.Windows.Forms.Label]::new()
    $wTitle.Text = if ($script:lang -eq "DE") { "SharePoint Einrichtung" } else { "SharePoint Setup" }; $wTitle.Font = [System.Drawing.Font]::new("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $wTitle.ForeColor = HC '#7C3AED'; $wTitle.BackColor = [System.Drawing.Color]::Transparent
    $wTitle.AutoSize = $true; $wTitle.Location = [System.Drawing.Point]::new(16, 13)
    $wHdr.Controls.Add($wTitle)

    $wSub = [System.Windows.Forms.Label]::new()
    $wSub.Text = $script:STR.WizSub
    $wSub.Font = $FS; $wSub.ForeColor = $wizLow
    $wSub.BackColor = [System.Drawing.Color]::Transparent; $wSub.AutoSize = $true
    $wSub.Location = [System.Drawing.Point]::new(190, 21)
    $wHdr.Controls.Add($wSub)

    $wSepHdr = [System.Windows.Forms.Panel]::new()
    $wSepHdr.SetBounds(0, 64, 660, 1); $wSepHdr.BackColor = $wizSep
    $wiz.Controls.Add($wSepHdr)

    # Tab buttons: Auto (Device Code) | Manual
    $wTabAuto = [System.Windows.Forms.Button]::new()
    $wTabAuto.Text = $script:STR.WizAutoTab; $wTabAuto.Size = [System.Drawing.Size]::new(220, 34)
    $wTabAuto.Location = [System.Drawing.Point]::new(16, 76); $wTabAuto.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $wTabAuto.Font = $FN; $wTabAuto.BackColor = HC '#7C3AED'; $wTabAuto.ForeColor = [System.Drawing.Color]::White
    $wTabAuto.Cursor = "Hand"; $wTabAuto.FlatAppearance.BorderColor = HC '#4C1D95'; $wTabAuto.FlatAppearance.BorderSize = 1
    $wiz.Controls.Add($wTabAuto)

    $wTabManual = [System.Windows.Forms.Button]::new()
    $wTabManual.Text = $script:STR.WizManualTab; $wTabManual.Size = [System.Drawing.Size]::new(220, 34)
    $wTabManual.Location = [System.Drawing.Point]::new(248, 76); $wTabManual.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $wTabManual.Font = $FN; $wTabManual.BackColor = HC '#374151'; $wTabManual.ForeColor = HC '#9CA3AF'
    $wTabManual.Cursor = "Hand"; $wTabManual.FlatAppearance.BorderColor = HC '#4B5563'; $wTabManual.FlatAppearance.BorderSize = 1
    $wiz.Controls.Add($wTabManual)

    $wSepTab = [System.Windows.Forms.Panel]::new()
    $wSepTab.SetBounds(0, 111, 660, 1); $wSepTab.BackColor = $wizSep
    $wiz.Controls.Add($wSepTab)

    # ── Auto panel ───────────────────────────────────────────
    $pAuto = [System.Windows.Forms.Panel]::new()
    $pAuto.SetBounds(0, 112, 660, 308); $pAuto.BackColor = $wizPnl
    $wiz.Controls.Add($pAuto)

    $aInfo = [System.Windows.Forms.Label]::new()
    $aInfo.Text = $script:STR.WizAutoInfo
    $aInfo.Font = $FN; $aInfo.ForeColor = $wizMid; $aInfo.BackColor = [System.Drawing.Color]::Transparent
    $aInfo.Size = [System.Drawing.Size]::new(630, 60); $aInfo.Location = [System.Drawing.Point]::new(20, 14)
    $pAuto.Controls.Add($aInfo)

    $aStep1 = [System.Windows.Forms.Label]::new()
    $aStep1.Text = $script:STR.WizStep1
    $aStep1.Font = $FB; $aStep1.ForeColor = $wizTxt; $aStep1.BackColor = [System.Drawing.Color]::Transparent
    $aStep1.AutoSize = $true; $aStep1.Location = [System.Drawing.Point]::new(20, 74)
    $pAuto.Controls.Add($aStep1)

    $aStep2 = [System.Windows.Forms.Label]::new()
    $aStep2.Text = $script:STR.WizStep2
    $aStep2.Font = $FB; $aStep2.ForeColor = $wizTxt; $aStep2.BackColor = [System.Drawing.Color]::Transparent
    $aStep2.AutoSize = $true; $aStep2.Location = [System.Drawing.Point]::new(20, 98)
    $pAuto.Controls.Add($aStep2)

    $aStep3 = [System.Windows.Forms.Label]::new()
    $aStep3.Text = $script:STR.WizStep3
    $aStep3.Font = $FB; $aStep3.ForeColor = $wizTxt; $aStep3.BackColor = [System.Drawing.Color]::Transparent
    $aStep3.AutoSize = $true; $aStep3.Location = [System.Drawing.Point]::new(20, 122)
    $pAuto.Controls.Add($aStep3)

    # Device code box
    $aCodeBox = [System.Windows.Forms.Panel]::new()
    $aCodeBoxBg = if ($script:lightMode) { HC "#EDE9FE" } else { HC "#1A1A2E" }
    $aCodeBox.SetBounds(20, 152, 620, 72); $aCodeBox.BackColor = $aCodeBoxBg
    $pAuto.Controls.Add($aCodeBox)

    $aCodeLabel = [System.Windows.Forms.Label]::new()
    $aCodeLabelInit = if ($script:lang -eq "DE") { "Klicke auf 'Gerätecode anfordern'" } else { "Click 'Get Device Code' to start" }
    $aCodeLabel.Text = $aCodeLabelInit
    $aCodeLabel.Font = [System.Drawing.Font]::new("Consolas", 14, [System.Drawing.FontStyle]::Bold)
    $aCodeLabelFg = if ($script:lightMode) { HC "#6D28D9" } else { HC "#FBBF24" }
    $aCodeLabel.ForeColor = $aCodeLabelFg
    $aCodeLabel.BackColor = [System.Drawing.Color]::Transparent
    $aCodeLabel.AutoSize = $true; $aCodeLabel.Location = [System.Drawing.Point]::new(16, 14)
    $aCodeBox.Controls.Add($aCodeLabel)

    $aUrlLabel = [System.Windows.Forms.Label]::new()
    $aUrlLabel.Text = ""; $aUrlLabel.Font = $FN
    $aUrlLabelFg = if ($script:lightMode) { HC '#1D4ED8' } else { HC '#93C5FD' }
    $aUrlLabel.ForeColor = $aUrlLabelFg; $aUrlLabel.BackColor = [System.Drawing.Color]::Transparent
    $aUrlLabel.Size = [System.Drawing.Size]::new(540, 20)
    $aUrlLabel.Location = [System.Drawing.Point]::new(16, 40)
    $aCodeBox.Controls.Add($aUrlLabel)

    $aStatusLabel = [System.Windows.Forms.Label]::new()
    $aStatusLabel.Text = ""; $aStatusLabel.Font = $FN
    $aStatusFg = if ($script:lightMode) { HC "#92400E" } else { HC "#FBBF24" }
    $aStatusLabel.ForeColor = $aStatusFg; $aStatusLabel.BackColor = [System.Drawing.Color]::Transparent
    $aStatusLabel.Size = [System.Drawing.Size]::new(620, 20); $aStatusLabel.Location = [System.Drawing.Point]::new(20, 226)
    $pAuto.Controls.Add($aStatusLabel)

    $aBtnGetCode = [System.Windows.Forms.Button]::new()
    $aBtnGetCode.Text = $script:STR.WizGetCode; $aBtnGetCode.Size = [System.Drawing.Size]::new(160, 34)
    $aBtnGetCode.Location = [System.Drawing.Point]::new(20, 254); $aBtnGetCode.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $aBtnGetCode.Font = $FB; $aBtnGetCode.BackColor = HC '#7C3AED'; $aBtnGetCode.ForeColor = [System.Drawing.Color]::White
    $aBtnGetCode.Cursor = "Hand"; $aBtnGetCode.FlatAppearance.BorderColor = HC '#4C1D95'; $aBtnGetCode.FlatAppearance.BorderSize = 1
    $aBtnGetCode.UseVisualStyleBackColor = $false
    $pAuto.Controls.Add($aBtnGetCode)

    $aBtnOpenBrowser = [System.Windows.Forms.Button]::new()
    $aBtnOpenBrowser.Text = $script:STR.WizOpenBrowser; $aBtnOpenBrowser.Size = [System.Drawing.Size]::new(240, 34)
    $aBtnOpenBrowser.Location = [System.Drawing.Point]::new(190, 254); $aBtnOpenBrowser.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $aBtnOpenBrowser.Font = $FN; $aBtnOpenBrowser.BackColor = HC '#1D4ED8'; $aBtnOpenBrowser.ForeColor = [System.Drawing.Color]::White
    $aBtnOpenBrowser.Cursor = "Hand"; $aBtnOpenBrowser.Enabled = $false
    $aBtnOpenBrowser.FlatAppearance.BorderColor = HC '#1E40AF'; $aBtnOpenBrowser.FlatAppearance.BorderSize = 1
    $aBtnOpenBrowser.UseVisualStyleBackColor = $false
    $pAuto.Controls.Add($aBtnOpenBrowser)

    # ── Manual panel ─────────────────────────────────────────
    $pManual = [System.Windows.Forms.Panel]::new()
    $pManual.SetBounds(0, 112, 660, 308); $pManual.BackColor = $wizPnl; $pManual.Visible = $false
    $wiz.Controls.Add($pManual)

    $mInfo = [System.Windows.Forms.Label]::new()
    $mInfo.Text = $script:STR.WizManualInfo
    $mInfo.Font = $FN; $mInfo.ForeColor = $wizMid; $mInfo.BackColor = [System.Drawing.Color]::Transparent
    $mInfo.Size = [System.Drawing.Size]::new(620, 42); $mInfo.Location = [System.Drawing.Point]::new(20, 14)
    $pManual.Controls.Add($mInfo)

    $mBtnDocs = [System.Windows.Forms.Label]::new()
    $mBtnDocs.Text      = $script:STR.WizOpenDocs
    $mBtnDocs.AutoSize  = $true
    $mBtnDocs.Location  = [System.Drawing.Point]::new(20, 52)
    $mBtnDocs.Font      = $FN
    $mDocsFg = if ($script:lightMode) { HC "#1D4ED8" } else { HC "#93C5FD" }
    $mBtnDocs.ForeColor = $mDocsFg
    $mBtnDocs.BackColor = [System.Drawing.Color]::Transparent
    $mBtnDocs.Cursor    = "Hand"
        $mBtnDocs.add_Click({ if ($script:lang -eq "DE") { Start-Process "https://github.com/C129H223N3O54/ACLens/blob/main/manuelle-sp-einrichtung.md" } else { Start-Process "https://github.com/C129H223N3O54/ACLens/blob/main/manual-sp-setup.md" } }) 
    $pManual.Controls.Add($mBtnDocs)

    function New-MField($label, $y, $placeholder) {
        $lbl = [System.Windows.Forms.Label]::new()
        $lbl.Text = $label; $lbl.Font = $FB
        $mfLblFg = if ($script:lightMode) { HC "#111111" } else { HC "#E2E8F0" }
        $lbl.ForeColor = $mfLblFg
        $lbl.BackColor = [System.Drawing.Color]::Transparent; $lbl.AutoSize = $true
        $lbl.Location = [System.Drawing.Point]::new(20, $y)
        $pManual.Controls.Add($lbl)
        $tb = [System.Windows.Forms.TextBox]::new()
        $tb.Size = [System.Drawing.Size]::new(618, 24); $tb.Location = [System.Drawing.Point]::new(20, $y + 22)
        $mfTbBg = if ($script:lightMode) { HC "#FFFFFF" } else { HC "#2D2D3F" }
        $mfTbFg = if ($script:lightMode) { HC "#111111" } else { HC "#9CA3AF" }
        $tb.BackColor = $mfTbBg; $tb.ForeColor = $mfTbFg
        $tb.Text = $placeholder; $tb.BorderStyle = "FixedSingle"; $tb.Font = $FN
        $tbPlaceholder = $placeholder
        $tb.Tag = $placeholder
        $tb.add_GotFocus({
            if ($this.Text -eq $this.Tag) {
                $this.Text = ""
                $mfTbFgActive = if ($script:lightMode) { HC "#111111" } else { HC "#E2E8F0" }
                $this.ForeColor = $mfTbFgActive
            }
        })
        $tb.add_LostFocus({
            if ([string]::IsNullOrWhiteSpace($this.Text)) {
                $this.Text = $this.Tag
                $mfTbFgReset = if ($script:lightMode) { HC "#6B7280" } else { HC "#9CA3AF" }
                $this.ForeColor = $mfTbFgReset
            }
        })
        $pManual.Controls.Add($tb)
        return $tb
    }

    $mTbTenant = New-MField $script:STR.WizTenantId     90  "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
    $mTbClient = New-MField $script:STR.WizClientId     140 "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
    $mTbSecret = New-MField $script:STR.WizSecret 190 "your-client-secret-value"
    $mTbSecret.UseSystemPasswordChar = $false

    $mBtnSave = [System.Windows.Forms.Button]::new()
    $mBtnSave.Text = $script:STR.WizSaveConnect; $mBtnSave.Size = [System.Drawing.Size]::new(160, 34)
    $mBtnSave.Location = [System.Drawing.Point]::new(20, 258); $mBtnSave.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $mBtnSave.Font = $FB; $mBtnSave.BackColor = HC '#7C3AED'; $mBtnSave.ForeColor = [System.Drawing.Color]::White
    $mBtnSave.Cursor = "Hand"; $mBtnSave.FlatAppearance.BorderColor = HC '#4C1D95'; $mBtnSave.FlatAppearance.BorderSize = 1
    $pManual.Controls.Add($mBtnSave)

    $mBtnSave.add_Click({
        $t = $mTbTenant.Text.Trim(); $c = $mTbClient.Text.Trim(); $sec = $mTbSecret.Text.Trim()
        if (-not $t -or -not $c -or -not $sec -or $t -like "*xxxx*") {
            [System.Windows.Forms.MessageBox]::Show($script:STR.WizMissingData, "Missing data", "OK", "Warning") | Out-Null
            return
        }
        Save-SPCredentials $t $c $sec
        Update-SPStatus
        [System.Windows.Forms.MessageBox]::Show($script:STR.WizSaved, "ACLens", "OK", "Information") | Out-Null
        $wiz.Close()
    })

    # ── Tab switching ─────────────────────────────────────────
    $wTabAuto.add_Click({
        $pAuto.Visible = $true; $pManual.Visible = $false
        $wTabAuto.BackColor = HC '#7C3AED'; $wTabAuto.ForeColor = [System.Drawing.Color]::White
        $wTabManual.BackColor = HC '#374151'; $wTabManual.ForeColor = HC '#9CA3AF'
    })
    $wTabManual.add_Click({
        $pManual.Visible = $true; $pAuto.Visible = $false
        $wTabManual.BackColor = HC '#7C3AED'; $wTabManual.ForeColor = [System.Drawing.Color]::White
        $wTabAuto.BackColor = HC '#374151'; $wTabAuto.ForeColor = HC '#9CA3AF'
    })

    # ── Device Code Flow ──────────────────────────────────────
    $script:deviceCode = $null
    $script:pollTimer  = $null

    $aBtnGetCode.add_Click({
        $aBtnGetCode.Enabled = $false
        $aStatusLabelInit = if ($script:lang -eq "DE") { "Gerätecode wird angefordert..." } else { "Requesting device code..." }
    $aStatusLabel.Text   = $aStatusLabelInit
        $aStatusFg = if ($script:lightMode) { HC "#92400E" } else { HC "#FBBF24" }
    $aStatusLabel.ForeColor = $aStatusFg
        $wiz.Refresh()

        try {
            # Request device code using well-known Microsoft first-party client
            # (Azure PowerShell App ID - widely trusted, allows app registration)
            $body = @{
                client_id = "1950a258-227b-4e31-a9cf-717495945fc2"
                scope     = "https://graph.microsoft.com/.default offline_access"
            }
            $resp = Invoke-RestMethod -Uri "https://login.microsoftonline.com/common/oauth2/v2.0/devicecode" `
                -Method Post -Body $body -ContentType "application/x-www-form-urlencoded"

            $script:deviceCode  = $resp
            $aCodeLabel.Text    = $resp.user_code
            $aUrlPrefix = if ($script:lang -eq "DE") { "Öffne: " } else { "Go to: " }
            $aUrlLabel.Text = $aUrlPrefix + $resp.verification_uri
            $aStatusLabel.Text  = ($script:STR.WizWaiting + " (" + $(if ($script:lang -eq "DE") { "läuft ab in " } else { "expires in " }) + [int]($resp.expires_in / 60) + $(if ($script:lang -eq "DE") { " Minuten)" } else { " minutes)" }))
            $aStatusLabel.ForeColor = HC '#FBBF24'
            $aBtnOpenBrowser.Enabled = $true

            # Start polling timer
            $script:pollTimer = [System.Windows.Forms.Timer]::new()
            $script:pollTimer.Interval = ($resp.interval * 1000)
            $script:pollTimer.add_Tick({
                try {
                    $tokenBody = @{
                        client_id   = "1950a258-227b-4e31-a9cf-717495945fc2"
                        grant_type  = "urn:ietf:params:oauth:grant-type:device_code"
                        device_code = $script:deviceCode.device_code
                    }
                    $token = Invoke-RestMethod -Uri "https://login.microsoftonline.com/common/oauth2/v2.0/token" `
                        -Method Post -Body $tokenBody -ContentType "application/x-www-form-urlencoded" -ErrorAction Stop

                    # Got token — now auto-register the app
                    $script:pollTimer.Stop()
                    $aStatusLabel.Text = $script:STR.WizCreating
                    $aStatusLabel.ForeColor = HC '#34D399'
                    $wiz.Refresh()

                    $headers = @{ Authorization = "Bearer $($token.access_token)"; "Content-Type" = "application/json" }

                    # Get tenant info
                    $me = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/me" -Headers $headers
                    $orgResp = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/organization" -Headers $headers
                    $tenantId = $orgResp.value[0].id

                    # Create App Registration
                    $appBody = @{
                        displayName    = "ACLens"
                        signInAudience = "AzureADMyOrg"
                        requiredResourceAccess = @(
                            @{
                                resourceAppId  = "00000003-0000-0000-c000-000000000000"  # Microsoft Graph
                                resourceAccess = @(
                                    @{ id = "9a5d68dd-52b0-4cc2-bd40-abcf44ac3a30"; type = "Role" }  # Application.Read.All
                                    @{ id = "0c0bf378-bf22-4279-b716-2d4f4c1e1be4"; type = "Role" }  # Sites.Read.All
                                    @{ id = "741f803b-c850-494e-b5df-cde7c675a1ca"; type = "Role" }  # User.Read.All
                                    @{ id = "98830695-27a2-44f7-8c18-0c3ebc9698f6"; type = "Role" }  # GroupMember.Read.All
                                )
                            }
                        )
                    } | ConvertTo-Json -Depth 5

                    $app = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/applications" `
                        -Method Post -Headers $headers -Body $appBody

                    # Create Service Principal
                    $spBody = @{ appId = $app.appId } | ConvertTo-Json
                    $sp = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/servicePrincipals" `
                        -Method Post -Headers $headers -Body $spBody

                    # Grant Admin Consent
                    $grantBody = @{
                        clientId   = $sp.id
                        consentType = "AllPrincipals"
                        principalId = $null
                        resourceId  = (Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId eq '00000003-0000-0000-c000-000000000000'" -Headers $headers).value[0].id
                        scope       = "Sites.Read.All User.Read.All GroupMember.Read.All"
                    } | ConvertTo-Json
                    Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/oauth2PermissionGrants" `
                        -Method Post -Headers $headers -Body $grantBody -ErrorAction SilentlyContinue | Out-Null

                    # Add Client Secret
                    $secretBody = @{
                        passwordCredential = @{
                            displayName = "ACLens Auto Secret"
                            endDateTime = (Get-Date).AddYears(2).ToString("o")
                        }
                    } | ConvertTo-Json
                    $secret = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/applications/$($app.id)/addPassword" `
                        -Method Post -Headers $headers -Body $secretBody

                    # Save credentials
                    Save-SPCredentials $tenantId $app.appId $secret.secretText
                    Update-SPStatus

                    $aStatusLabel.Text = "✔  $($script:STR.WizDone) Client ID: $($app.appId)"
                    $aStatusLabel.ForeColor = HC '#34D399'
                    $aCodeLabel.Text = $script:STR.WizDone
                    $aCodeLabel.ForeColor = HC '#34D399'

                    Start-Sleep -Milliseconds 1500
                    $wiz.Close()

                } catch {
                    $err = $_.Exception.Message
                    if ($err -notlike "*authorization_pending*" -and $err -notlike "*slow_down*") {
                        if ($err -like "*authorization_declined*" -or $err -like "*expired*") {
                            $script:pollTimer.Stop()
                            $aStatusLabel.Text = $script:STR.WizCancelled
                            $aStatusLabel.ForeColor = HC '#F87171'
                            $aBtnGetCode.Enabled = $true
                        }
                    }
                }
            })
            $script:pollTimer.Start()

        } catch {
            $aStatusLabel.Text = "Error: $($_.Exception.Message)"
            $aStatusLabel.ForeColor = HC '#F87171'
            $aBtnGetCode.Enabled = $true
        }
    })

    $aBtnOpenBrowser.add_Click({
        Start-Process "https://microsoft.com/devicelogin"
    })

    # Footer separator + close
    $wSepFtr = [System.Windows.Forms.Panel]::new()
    $wSepFtr.SetBounds(0, 420, 660, 1); $wSepFtr.BackColor = $wizSep
    $wiz.Controls.Add($wSepFtr)

    $wBtnClose = [System.Windows.Forms.Button]::new()
    $wBtnClose.Text = $script:STR.WizClose; $wBtnClose.Size = [System.Drawing.Size]::new(100, 34)
    $wBtnClose.Location = [System.Drawing.Point]::new(544, 434); $wBtnClose.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $wBtnClose.Font = $FN; $wBtnCloseBg = if ($script:lightMode) { HC '#E0E0E0' } else { HC '#374151' }
    $wBtnCloseFg = if ($script:lightMode) { HC '#111111' } else { HC '#E2E8F0' }
    $wBtnClose.BackColor = $wBtnCloseBg; $wBtnClose.ForeColor = $wBtnCloseFg
    $wBtnClose.UseVisualStyleBackColor = $false
    $wBtnClose.Cursor = "Hand"; $wBtnClose.FlatAppearance.BorderColor = HC '#4B5563'; $wBtnClose.FlatAppearance.BorderSize = 1
    $wBtnClose.add_Click({
        if ($null -ne $script:pollTimer) { $script:pollTimer.Stop() }
        $wiz.Close()
    })
    $wiz.Controls.Add($wBtnClose)

    $wiz.add_FormClosing({
        if ($null -ne $script:pollTimer) { $script:pollTimer.Stop() }
    })

    $wiz.ShowDialog($form) | Out-Null
}

$btnOpenLast.add_Click({
    if ($script:lastReport -and (Test-Path $script:lastReport)) {
        Start-Process $script:lastReport
    }
})

$changelogText = @"
ACLens Changelog
================

v0.4.0-beta  (2026-03-27)
---------------------------
- Fixed: switch statement spacing causing ParseException on PS 5.1
- Version label added to main window header

v0.2.0-alpha  (2026-03-27)
---------------------------
- SharePoint Online support via Microsoft Graph API
- NTFS/SharePoint tab switcher in main GUI
- Setup Wizard: automatic App Registration via Device Code Flow
- Setup Wizard: manual mode with Tenant ID / Client ID / Secret
- Encrypted credential storage (sp_credentials.xml, DPAPI)
- SP Scanner: Sites, Libraries, Folders, Groups & Members, External Sharing
- SP HTML report with collapsible library/folder tree
- SP JSON snapshot saved alongside every SP report
- Compare Scans extended: NTFS Compare + SharePoint Compare tabs
- SP Diff report: added/removed folders, changed permissions

v0.1.0-alpha  (2026-03-14)
---------------------------
- Initial public alpha release
- Recursive NTFS permission scan with configurable depth
- Detects permission changes vs. parent folder
- Interactive HTML report (filter, search, expand/collapse, print)
- Collapsible legend in HTML report
- Automatic JSON snapshot export for baseline comparison
- Compare Scans: diff two scan results (added/removed folders, changed permissions)
- GUI with progress bar, status and cancel support
- Open Last Report / Changelog / Compare Scans buttons
- Admin check dialog on startup with restart-as-admin option
- Output saved to script folder by default
- Compatible: Windows 10/11, Server 2016-2025, PowerShell 5.1+
"@

$btnChangelog.add_Click({
    $dlg = [System.Windows.Forms.Form]::new()
    $dlg.Text            = "ACLens - Changelog"
    $dlg.Size            = [System.Drawing.Size]::new(620, 520)
    $dlg.MinimumSize     = [System.Drawing.Size]::new(420, 300)
    $dlg.StartPosition   = "CenterParent"
    $dlg.FormBorderStyle = "Sizable"
    $dlg.MaximizeBox     = $true
    $dlgBg  = if ($script:lightMode) { HC "#F5F5F5" } else { HC "#1A1A2E" }
    $logBg  = if ($script:lightMode) { HC "#FFFFFF" } else { HC "#1A1A2E" }
    $logFg  = if ($script:lightMode) { HC "#111111" } else { HC "#E2E8F0" }
    $dlg.BackColor = $dlgBg

    $txtLog = [System.Windows.Forms.RichTextBox]::new()
    $txtLog.Dock        = "Fill"
    $txtLog.ScrollBars  = "Vertical"
    $txtLog.BackColor   = $logBg
    $txtLog.ForeColor   = $logFg
    $txtLog.Font        = [System.Drawing.Font]::new("Consolas", 9)
    $txtLog.BorderStyle = "None"
    $txtLog.Padding     = [System.Windows.Forms.Padding]::new(12)
    $txtLog.Text        = $changelogText
    $txtLog.ReadOnly    = $true
    # Re-apply colors after Text assignment (WinForms may reset them)
    $txtLog.BackColor   = $logBg
    $txtLog.ForeColor   = $logFg
    $dlg.Controls.Add($txtLog)

    $dlg.ShowDialog($form) | Out-Null
})


# ── Compare Scan Helper Functions ────────────────────────────
function Load-JsonSnapshot {
    param([string]$Path)
    $raw = Get-Content -Path $Path -Raw -Encoding UTF8 | ConvertFrom-Json
    $result = @{}
    foreach ($item in $raw) {
        $rules = @()
        foreach ($r in $item.Rules) {
            $rules += [PSCustomObject]@{
                Identity         = $r.Identity
                AccessType       = $r.AccessType
                Rights           = $r.Rights
                IsInherited      = $r.IsInherited
                InheritanceFlags = $r.InheritanceFlags
                PropagationFlags = $r.PropagationFlags
            }
        }
        $result[$item.Path] = [PSCustomObject]@{
            Path                    = $item.Path
            Owner                   = $item.Owner
            InheritanceEnabled      = $item.InheritanceEnabled
            AreAccessRulesProtected = $item.AreAccessRulesProtected
            Rules                   = $rules
        }
    }
    return $result
}

function Get-RulesKey {
    param([array]$Rules)
    ($Rules | Sort-Object Identity, AccessType, Rights, InheritanceFlags, PropagationFlags |
        ForEach-Object { "$($_.Identity)|$($_.AccessType)|$($_.Rights)|$($_.IsInherited)|$($_.InheritanceFlags)|$($_.PropagationFlags)" }) -join ";"
}

# ── Compare Window ─────────────────────────────────────────
$btnCompare.add_Click({

    $cWin = [System.Windows.Forms.Form]::new()
    $cWin.Text            = "ACLens - Compare Scans"
    $cWin.Size            = [System.Drawing.Size]::new(700, 480)
    $cWin.StartPosition   = "CenterParent"
    $cWin.FormBorderStyle = "FixedDialog"
    $cWin.MaximizeBox     = $false
    $cWin.BackColor       = HC '#1E1E2E'
    $cWin.Font            = $FN

    # Header
    $cHdr = [System.Windows.Forms.Panel]::new()
    $cHdr.Size = [System.Drawing.Size]::new(700, 52); $cHdr.Location = [System.Drawing.Point]::new(0, 0)
    $cHdr.BackColor = HC '#1A1A2E'
    $cWin.Controls.Add($cHdr)

    $cTitle = [System.Windows.Forms.Label]::new()
    $cTitle.Text = "Compare Scans"; $cTitle.Font = [System.Drawing.Font]::new("Segoe UI", 13, [System.Drawing.FontStyle]::Bold)
    $cTitle.ForeColor = HC '#A78BFA'; $cTitle.BackColor = [System.Drawing.Color]::Transparent
    $cTitle.AutoSize = $true; $cTitle.Location = [System.Drawing.Point]::new(16, 13)
    $cHdr.Controls.Add($cTitle)

    $cSub = [System.Windows.Forms.Label]::new()
    $cSub.Font = $FS; $cSub.ForeColor = HC '#6B7280'; $cSub.BackColor = [System.Drawing.Color]::Transparent
    $cSub.AutoSize = $true; $cSub.Location = [System.Drawing.Point]::new(180, 20)
    $cHdr.Controls.Add($cSub)

    $cSepHdr = [System.Windows.Forms.Panel]::new()
    $cSepHdr.SetBounds(0, 52, 700, 1); $cSepHdr.BackColor = HC '#4B5563'
    $cWin.Controls.Add($cSepHdr)

    # ── Mode selector tabs ────────────────────────────────────
    $cTabNTFS = [System.Windows.Forms.Button]::new()
    $cTabNTFS.Text = "  NTFS Compare"; $cTabNTFS.Size = [System.Drawing.Size]::new(150, 34)
    $cTabNTFS.Location = [System.Drawing.Point]::new(16, 62); $cTabNTFS.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $cTabNTFS.Font = $FB; $cTabNTFS.BackColor = HC '#7C3AED'; $cTabNTFS.ForeColor = [System.Drawing.Color]::White
    $cTabNTFS.Cursor = "Hand"; $cTabNTFS.FlatAppearance.BorderColor = HC '#4C1D95'; $cTabNTFS.FlatAppearance.BorderSize = 1
    $cWin.Controls.Add($cTabNTFS)

    $cTabSP = [System.Windows.Forms.Button]::new()
    $cTabSP.Text = "  SharePoint Compare"; $cTabSP.Size = [System.Drawing.Size]::new(180, 34)
    $cTabSP.Location = [System.Drawing.Point]::new(174, 62); $cTabSP.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $cTabSP.Font = $FB; $cTabSP.BackColor = HC '#374151'; $cTabSP.ForeColor = HC '#9CA3AF'
    $cTabSP.Cursor = "Hand"; $cTabSP.FlatAppearance.BorderColor = HC '#4B5563'; $cTabSP.FlatAppearance.BorderSize = 1
    $cWin.Controls.Add($cTabSP)

    $cSepTab = [System.Windows.Forms.Panel]::new()
    $cSepTab.SetBounds(0, 97, 700, 1); $cSepTab.BackColor = HC '#374151'
    $cWin.Controls.Add($cSepTab)

    # ── NTFS Compare panel ────────────────────────────────────
    $pNTFS = [System.Windows.Forms.Panel]::new()
    $pNTFS.SetBounds(0, 98, 700, 280); $pNTFS.BackColor = HC '#1E1E2E'
    $cWin.Controls.Add($pNTFS)

    # Baseline JSON
    $nLbl1 = [System.Windows.Forms.Label]::new()
    $nLbl1.Text = "Baseline (JSON):"; $nLbl1.Font = $FB; $nLbl1.ForeColor = HC '#E2E8F0'
    $nLbl1.BackColor = [System.Drawing.Color]::Transparent; $nLbl1.AutoSize = $true
    $nLbl1.Location = [System.Drawing.Point]::new(20, 14)
    $pNTFS.Controls.Add($nLbl1)

    $nTbJson = [System.Windows.Forms.TextBox]::new()
    $nTbJson.Size = [System.Drawing.Size]::new(528, 26); $nTbJson.Location = [System.Drawing.Point]::new(20, 36)
    $nTbJson.BackColor = HC '#2D2D3F'; $nTbJson.ForeColor = HC '#E2E8F0'; $nTbJson.BorderStyle = "FixedSingle"; $nTbJson.Font = $FN
    if ($script:lastJson -and (Test-Path $script:lastJson)) { $nTbJson.Text = $script:lastJson }
    $pNTFS.Controls.Add($nTbJson)

    $nBtnJson = [System.Windows.Forms.Button]::new()
    $nBtnJson.Text = "Browse..."; $nBtnJson.Size = [System.Drawing.Size]::new(110, 26)
    $nBtnJson.Location = [System.Drawing.Point]::new(554, 36); $nBtnJson.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $nBtnJson.BackColor = HC '#374151'; $nBtnJson.ForeColor = HC '#E2E8F0'; $nBtnJson.UseVisualStyleBackColor = $false; $nBtnJson.Cursor = "Hand"
    $nBtnJson.FlatAppearance.BorderColor = HC '#4B5563'; $nBtnJson.FlatAppearance.BorderSize = 1
    $nBtnJson.add_Click({
        $ofd = [System.Windows.Forms.OpenFileDialog]::new()
        $ofd.Filter = "ACLens JSON (*.json)|*.json"
        if ($ofd.ShowDialog() -eq "OK") { $nTbJson.Text = $ofd.FileName }
    })
    $pNTFS.Controls.Add($nBtnJson)

    # Current scan path
    $nLbl2 = [System.Windows.Forms.Label]::new()
    $nLbl2.Text = "Current Scan Path:"; $nLbl2.Font = $FB; $nLbl2.ForeColor = HC '#E2E8F0'
    $nLbl2.BackColor = [System.Drawing.Color]::Transparent; $nLbl2.AutoSize = $true
    $nLbl2.Location = [System.Drawing.Point]::new(20, 74)
    $pNTFS.Controls.Add($nLbl2)

    $nTbPath = [System.Windows.Forms.TextBox]::new()
    $nTbPath.Size = [System.Drawing.Size]::new(528, 26); $nTbPath.Location = [System.Drawing.Point]::new(20, 96)
    $nTbPath.BackColor = HC '#2D2D3F'; $nTbPath.ForeColor = HC '#E2E8F0'; $nTbPath.BorderStyle = "FixedSingle"; $nTbPath.Font = $FN
    if ($tbPath.Text) { $nTbPath.Text = $tbPath.Text }
    $pNTFS.Controls.Add($nTbPath)

    $nBtnPath = [System.Windows.Forms.Button]::new()
    $nBtnPath.Text = "Browse..."; $nBtnPath.Size = [System.Drawing.Size]::new(110, 26)
    $nBtnPath.Location = [System.Drawing.Point]::new(554, 96); $nBtnPath.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $nBtnPath.BackColor = HC '#374151'; $nBtnPath.ForeColor = HC '#E2E8F0'; $nBtnPath.UseVisualStyleBackColor = $false; $nBtnPath.Cursor = "Hand"
    $nBtnPath.FlatAppearance.BorderColor = HC '#4B5563'; $nBtnPath.FlatAppearance.BorderSize = 1
    $nBtnPath.add_Click({
        $fbd = [System.Windows.Forms.FolderBrowserDialog]::new()
        if ($nTbPath.Text -and (Test-Path $nTbPath.Text)) { $fbd.SelectedPath = $nTbPath.Text }
        if ($fbd.ShowDialog() -eq "OK") { $nTbPath.Text = $fbd.SelectedPath }
    })
    $pNTFS.Controls.Add($nBtnPath)

    # ── SP Compare panel ──────────────────────────────────────
    $pSP = [System.Windows.Forms.Panel]::new()
    $pSP.SetBounds(0, 98, 700, 280); $pSP.BackColor = HC '#1E1E2E'; $pSP.Visible = $false
    $cWin.Controls.Add($pSP)

    $sLbl1 = [System.Windows.Forms.Label]::new()
    $sLbl1.Text = "Baseline SP Snapshot (JSON):"; $sLbl1.Font = $FB; $sLbl1.ForeColor = HC '#E2E8F0'
    $sLbl1.BackColor = [System.Drawing.Color]::Transparent; $sLbl1.AutoSize = $true
    $sLbl1.Location = [System.Drawing.Point]::new(20, 14)
    $pSP.Controls.Add($sLbl1)

    $sTbJson = [System.Windows.Forms.TextBox]::new()
    $sTbJson.Size = [System.Drawing.Size]::new(528, 26); $sTbJson.Location = [System.Drawing.Point]::new(20, 36)
    $sTbJson.BackColor = HC '#2D2D3F'; $sTbJson.ForeColor = HC '#E2E8F0'; $sTbJson.BorderStyle = "FixedSingle"; $sTbJson.Font = $FN
    $pSP.Controls.Add($sTbJson)

    $sBtnJson = [System.Windows.Forms.Button]::new()
    $sBtnJson.Text = "Browse..."; $sBtnJson.Size = [System.Drawing.Size]::new(110, 26)
    $sBtnJson.Location = [System.Drawing.Point]::new(554, 36); $sBtnJson.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $sBtnJson.BackColor = HC '#374151'; $sBtnJson.ForeColor = HC '#E2E8F0'; $sBtnJson.UseVisualStyleBackColor = $false; $sBtnJson.Cursor = "Hand"
    $sBtnJson.FlatAppearance.BorderColor = HC '#4B5563'; $sBtnJson.FlatAppearance.BorderSize = 1
    $sBtnJson.add_Click({
        $ofd = [System.Windows.Forms.OpenFileDialog]::new()
        $ofd.Filter = "ACLens SP Snapshot (*.json)|*.json"
        if ($ofd.ShowDialog() -eq "OK") { $sTbJson.Text = $ofd.FileName }
    })
    $pSP.Controls.Add($sBtnJson)

    $sLbl2 = [System.Windows.Forms.Label]::new()
    $sLbl2.Text = "Current SharePoint Site URL:"; $sLbl2.Font = $FB; $sLbl2.ForeColor = HC '#E2E8F0'
    $sLbl2.BackColor = [System.Drawing.Color]::Transparent; $sLbl2.AutoSize = $true
    $sLbl2.Location = [System.Drawing.Point]::new(20, 74)
    $pSP.Controls.Add($sLbl2)

    $sTbUrl = [System.Windows.Forms.TextBox]::new()
    $sTbUrl.Size = [System.Drawing.Size]::new(640, 26); $sTbUrl.Location = [System.Drawing.Point]::new(20, 96)
    $sTbUrl.BackColor = HC '#2D2D3F'; $sTbUrl.ForeColor = HC '#6B7280'; $sTbUrl.BorderStyle = "FixedSingle"; $sTbUrl.Font = $FN
    $sTbUrl.Text = if ($spTbUrl.Text -notlike "*yourtenant*") { $spTbUrl.Text } else { "https://yourtenant.sharepoint.com/sites/yoursite" }
    $pSP.Controls.Add($sTbUrl)

    $sNote = [System.Windows.Forms.Label]::new()
    $sNote.Text = "Note: A new live SP scan will run and be compared against the baseline JSON."
    $sNote.Font = $FS; $sNote.ForeColor = HC '#6B7280'; $sNote.BackColor = [System.Drawing.Color]::Transparent
    $sNote.Size = [System.Drawing.Size]::new(640, 18); $sNote.Location = [System.Drawing.Point]::new(20, 130)
    $pSP.Controls.Add($sNote)

    # ── Tab switching ─────────────────────────────────────────
    $script:cMode = "NTFS"
    $cTabNTFS.add_Click({
        $script:cMode = "NTFS"
        $pNTFS.Visible = $true;  $pSP.Visible = $false
        $cTabNTFS.BackColor = HC '#7C3AED'; $cTabNTFS.ForeColor = [System.Drawing.Color]::White
        $cTabSP.BackColor   = HC '#374151'; $cTabSP.ForeColor   = HC '#9CA3AF'
        $cSub.Text = "NTFS permission diff"
    })
    $cTabSP.add_Click({
        $script:cMode = "SP"
        $pSP.Visible = $true;  $pNTFS.Visible = $false
        $cTabSP.BackColor   = HC '#7C3AED'; $cTabSP.ForeColor   = [System.Drawing.Color]::White
        $cTabNTFS.BackColor = HC '#374151'; $cTabNTFS.ForeColor = HC '#9CA3AF'
        $cSub.Text = "SharePoint permission diff"
    })
    $cSub.Text = "NTFS permission diff"

    # ── Shared: Status + Progress ─────────────────────────────
    $cSepMid = [System.Windows.Forms.Panel]::new()
    $cSepMid.SetBounds(0, 378, 700, 1); $cSepMid.BackColor = HC '#4B5563'
    $cWin.Controls.Add($cSepMid)

    $cStatus = [System.Windows.Forms.Label]::new()
    $cStatus.Text = "Ready."; $cStatus.Font = $FN; $cStatus.ForeColor = HC '#6B7280'
    $cStatus.BackColor = [System.Drawing.Color]::Transparent
    $cStatus.Size = [System.Drawing.Size]::new(660, 18); $cStatus.Location = [System.Drawing.Point]::new(20, 390)
    $cWin.Controls.Add($cStatus)

    $cBar = [System.Windows.Forms.ProgressBar]::new()
    $cBar.Size = [System.Drawing.Size]::new(660, 12); $cBar.Location = [System.Drawing.Point]::new(20, 412)
    $cBar.Style = "Continuous"; $cBar.Minimum = 0; $cBar.Maximum = 100
    $cWin.Controls.Add($cBar)

    # ── Footer ────────────────────────────────────────────────
    $cSepFtr = [System.Windows.Forms.Panel]::new()
    $cSepFtr.SetBounds(0, 430, 700, 1); $cSepFtr.BackColor = HC '#4B5563'
    $cWin.Controls.Add($cSepFtr)

    $btnRun = [System.Windows.Forms.Button]::new()
    $btnRun.Text = "Run Comparison"; $btnRun.Size = [System.Drawing.Size]::new(160, 36)
    $btnRun.Location = [System.Drawing.Point]::new(16, 440); $btnRun.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnRun.Font = $FB; $btnRun.BackColor = HC '#7C3AED'; $btnRun.ForeColor = [System.Drawing.Color]::White
    $btnRun.Cursor = "Hand"; $btnRun.FlatAppearance.BorderColor = HC '#4C1D95'; $btnRun.FlatAppearance.BorderSize = 1
    $cWin.Controls.Add($btnRun)

    $btnRun.add_Click({
        $btnRun.Enabled = $false
        $cBar.Value     = 0
        $cStatus.ForeColor = HC '#9CA3AF'

        if ($script:cMode -eq "NTFS") {
            # ── NTFS diff ─────────────────────────────────────
            $jsonPath = $nTbJson.Text.Trim()
            $scanPath = $nTbPath.Text.Trim()
            if (-not $jsonPath -or -not (Test-Path $jsonPath)) {
                [System.Windows.Forms.MessageBox]::Show("Please select a valid baseline JSON file.", "Missing input", "OK", "Warning") | Out-Null
                $btnRun.Enabled = $true; return
            }
            if (-not $scanPath -or -not (Test-Path $scanPath -PathType Container)) {
                [System.Windows.Forms.MessageBox]::Show("Please enter a valid scan path.", "Missing input", "OK", "Warning") | Out-Null
                $btnRun.Enabled = $true; return
            }
            try {
                $cStatus.Text = "Loading baseline..."; $cWin.Refresh()
                $oldData  = Load-JsonSnapshot -Path $jsonPath

                $cStatus.Text = "Step 1/2: Scanning folders..."; $cWin.Refresh()
                $allFolders = Get-AllFolders -RootPath $scanPath -MaxDepth 0
                $total = $allFolders.Count
                $newData = @{}
                $rootClean = $scanPath.TrimEnd('\')
                $counter   = 0

                foreach ($fp in $allFolders) {
                    $counter++
                    $pct = [int](($counter / $total) * 100)
                    if ($cBar.Value -ne $pct) { $cBar.Value = $pct }
                    if ($counter % 5 -eq 0 -or $counter -eq $total) {
                        $short = if ($fp.Length -gt 70) { "..." + $fp.Substring($fp.Length-67) } else { $fp }
                        $cStatus.Text = "Step 2/2: $counter / $total  —  $short"; $cWin.Refresh()
                    }
                    $aclData = Get-FolderACL -FolderPath $fp
                    $newData[$fp] = [PSCustomObject]@{
                        Path = $fp; Owner = $aclData.Owner
                        InheritanceEnabled = $aclData.InheritanceEnabled
                        AreAccessRulesProtected = $aclData.AreAccessRulesProtected
                        Rules = $aclData.Rules
                    }
                }

                $cBar.Value = 99; $cStatus.Text = "Generating diff report..."; $cWin.Refresh()
                $ts      = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
                $outPath = Join-Path $PSScriptRoot "ACLens_Compare_$ts.html"
                $result  = New-CompareHTMLReport -OldData $oldData -NewData $newData `
                    -OldLabel $jsonPath -NewLabel $scanPath -OutputPath $outPath

                $cBar.Value = 100
                $cStatus.Text = "Done!  Changed: $($result.Changed)  Added: $($result.Added)  Removed: $($result.Removed)"
                $cStatus.ForeColor = HC '#34D399'
                Start-Process $outPath

            } catch {
                $cStatus.Text = "Error: $($_.Exception.Message)"; $cStatus.ForeColor = HC '#F87171'
            }

        } else {
            # ── SharePoint diff ───────────────────────────────
            $jsonPath = $sTbJson.Text.Trim()
            $siteUrl  = $sTbUrl.Text.Trim()

            if (-not $jsonPath -or -not (Test-Path $jsonPath)) {
                [System.Windows.Forms.MessageBox]::Show("Please select a valid baseline SP JSON file.", "Missing input", "OK", "Warning") | Out-Null
                $btnRun.Enabled = $true; return
            }
            if (-not $siteUrl -or $siteUrl -like "*yourtenant*") {
                [System.Windows.Forms.MessageBox]::Show("Please enter a valid SharePoint Site URL.", "Missing input", "OK", "Warning") | Out-Null
                $btnRun.Enabled = $true; return
            }
            $creds = Get-SPCredentials
            if (-not $creds.IsConfigured) {
                [System.Windows.Forms.MessageBox]::Show("SharePoint is not configured. Use Setup / Reconnect first.", "Not configured", "OK", "Warning") | Out-Null
                $btnRun.Enabled = $true; return
            }

            try {
                $cStatus.Text = "Loading SP baseline..."; $cWin.Refresh()
                $baselineRaw  = Get-Content $jsonPath -Raw | ConvertFrom-Json
                $oldSPData    = @{}
                foreach ($item in $baselineRaw) {
                    $oldSPData[$item.FullPath] = $item
                }

                $cStatus.Text = "Running live SP scan..."; $cWin.Refresh()
                $ts      = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
                $tmpOut  = Join-Path $env:TEMP "ACLens_SPtemp_$ts.html"
                $tmpJson = Join-Path $PSScriptRoot "ACLens_SP_Report_$ts.json"

                Start-SPScan -SiteUrl $siteUrl -MaxDepth 0 -OutputPath $tmpOut `
                    -StatusLabel $cStatus -ProgressBar $cBar -StatsLabel ([System.Windows.Forms.Label]::new()) -FormRef $cWin

                # Load new scan data from JSON snapshot
                if (Test-Path $tmpJson) {
                    $newRaw    = Get-Content $tmpJson -Raw | ConvertFrom-Json
                    $newSPData = @{}
                    foreach ($item in $newRaw) { $newSPData[$item.FullPath] = $item }

                    $cStatus.Text = "Generating SP diff report..."; $cWin.Refresh()
                    $outPath = Join-Path $PSScriptRoot "ACLens_SP_Compare_$ts.html"
                    New-SPCompareHTMLReport -OldData $oldSPData -NewData $newSPData `
                        -OldLabel $jsonPath -NewLabel $siteUrl -OutputPath $outPath

                    $cBar.Value = 100
                    $cStatus.Text = "Done!  SP diff report saved."
                    $cStatus.ForeColor = HC '#34D399'
                    Start-Process $outPath
                } else {
                    throw "SP scan snapshot not found. Scan may have failed."
                }

            } catch {
                $cStatus.Text = "Error: $($_.Exception.Message)"; $cStatus.ForeColor = HC '#F87171'
            }
        }

        $btnRun.Enabled = $true
    })

    $cWin.ShowDialog($form) | Out-Null
})



    [System.Windows.Forms.Application]::Run($form)

} # end Start-ACLensGUI



# ── Launch ───────────────────────────────────────────────────
Start-ACLensGUI
# ── SharePoint HTML Report ────────────────────────────────────
function New-SPHTMLReport {
    param(
        [string]$SiteUrl,
        [string]$SiteId,
        [string]$SiteName,
        [array]$SitePerms,
        [array]$Groups,
        [string]$ExternalSharing,
        [array]$Drives,
        [string]$OutputPath
    )

    $reportDate   = Get-Date -Format "dd.MM.yyyy HH:mm:ss"
    $computerName = $env:COMPUTERNAME
    $currentUser  = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name

    $extClass = switch ($ExternalSharing) {
        "Disabled"                 { "ext-disabled" }
        "Existing external users only" { "ext-limited" }
        default                    { "ext-open" }
    }

    # ── Site permissions table ────────────────────────────────
    $sitePermRows = [System.Text.StringBuilder]::new()
    foreach ($p in $SitePerms) {
        [void]$sitePermRows.Append("<tr><td class='sp-identity'>$([System.Web.HttpUtility]::HtmlEncode($p.GrantedTo))</td>")
        [void]$sitePermRows.Append("<td><span class='sp-type-badge'>$([System.Web.HttpUtility]::HtmlEncode($p.Type))</span></td>")
        [void]$sitePermRows.Append("<td class='sp-roles'>$([System.Web.HttpUtility]::HtmlEncode($p.Roles))</td></tr>")
    }

    # ── Groups table ──────────────────────────────────────────
    $groupRows = [System.Text.StringBuilder]::new()
    foreach ($g in $Groups) {
        [void]$groupRows.Append("<tr><td class='sp-identity'>$([System.Web.HttpUtility]::HtmlEncode($g.Name))</td>")
        [void]$groupRows.Append("<td class='sp-count'>$($g.Count)</td>")
        [void]$groupRows.Append("<td class='sp-members'>$([System.Web.HttpUtility]::HtmlEncode($g.Members))</td></tr>")
    }

    # ── Library + folder rows ─────────────────────────────────
    $libRows = [System.Text.StringBuilder]::new()
    $rowIdx  = 0
    foreach ($drive in $Drives) {
        $rowIdx++
        $libPermsHtml = ""
        if ($drive.Permissions.Count -gt 0) {
            $libPermsHtml = "<table class='sp-mini-table'><thead><tr><th>Principal</th><th>Type</th><th>Roles</th></tr></thead><tbody>"
            foreach ($p in $drive.Permissions) {
                $libPermsHtml += "<tr><td>$([System.Web.HttpUtility]::HtmlEncode($p.GrantedTo))</td><td>$([System.Web.HttpUtility]::HtmlEncode($p.Type))</td><td>$([System.Web.HttpUtility]::HtmlEncode($p.Roles))</td></tr>"
            }
            $libPermsHtml += "</tbody></table>"
        } else {
            $libPermsHtml = "<span class='sp-note'>No unique permissions</span>"
        }

        [void]$libRows.Append("<div class='sp-library' id='lib-$rowIdx'>")
        [void]$libRows.Append("<div class='sp-lib-header' onclick='toggleSP(""lib-body-$rowIdx"",this)'>")
        [void]$libRows.Append("<span class='sp-toggle'>&#9658;</span>")
        [void]$libRows.Append("<span class='sp-lib-icon'>&#128218;</span>")
        [void]$libRows.Append("<span class='sp-lib-name'>$([System.Web.HttpUtility]::HtmlEncode($drive.Name))</span>")
        [void]$libRows.Append("<span class='sp-lib-type'>$([System.Web.HttpUtility]::HtmlEncode($drive.DriveType))</span>")
        [void]$libRows.Append("<span class='sp-folder-count'>$($drive.Folders.Count) folders</span>")
        [void]$libRows.Append("</div>")
        [void]$libRows.Append("<div class='sp-lib-body' id='lib-body-$rowIdx'>")
        [void]$libRows.Append("<div class='sp-section-title'>Library Permissions</div>")
        [void]$libRows.Append($libPermsHtml)

        if ($drive.Folders.Count -gt 0) {
            [void]$libRows.Append("<div class='sp-section-title' style='margin-top:12px'>Folder Permissions</div>")
            [void]$libRows.Append("<div class='sp-folder-list'>")
            foreach ($folder in $drive.Folders) {
                $indent = "&nbsp;" * ($folder.Depth * 4)
                $statusClass = if ($folder.HasUniquePerms) { "sp-folder-unique" } elseif ($folder.Error) { "sp-folder-error" } else { "sp-folder-inherited" }
                $statusBadge = if ($folder.HasUniquePerms) { "<span class='sp-badge sp-badge-unique'>Unique</span>" } elseif ($folder.Error) { "<span class='sp-badge sp-badge-error'>Error</span>" } else { "<span class='sp-badge sp-badge-inherited'>Inherited</span>" }

                $fRowIdx = "${rowIdx}_$($drive.Folders.IndexOf($folder))"
                [void]$libRows.Append("<div class='sp-folder-row $statusClass'>")
                [void]$libRows.Append("<div class='sp-folder-hdr' onclick='toggleSP(""fp-$fRowIdx"",this)'>")
                [void]$libRows.Append("<span class='sp-toggle sp-toggle-sm'>&#9658;</span>")
                [void]$libRows.Append("<span class='sp-folder-icon'>&#128193;</span>")
                [void]$libRows.Append("$indent<span class='sp-folder-name'>$([System.Web.HttpUtility]::HtmlEncode($folder.Path))</span>")
                [void]$libRows.Append($statusBadge)
                [void]$libRows.Append("</div>")

                if ($folder.HasUniquePerms -and $folder.Permissions.Count -gt 0) {
                    [void]$libRows.Append("<div class='sp-folder-detail' id='fp-$fRowIdx'>")
                    [void]$libRows.Append("<table class='sp-mini-table'><thead><tr><th>Principal</th><th>Roles</th><th>Inherited</th><th>Link</th></tr></thead><tbody>")
                    foreach ($p in $folder.Permissions) {
                        if (-not $p.IsInherited) {
                            $linkBadge = if ($p.Link) { "<span class='sp-badge sp-badge-link'>Sharing Link</span>" } else { "" }
                            [void]$libRows.Append("<tr><td>$([System.Web.HttpUtility]::HtmlEncode($p.GrantedTo))</td><td>$([System.Web.HttpUtility]::HtmlEncode($p.Roles))</td><td>No</td><td>$linkBadge</td></tr>")
                        }
                    }
                    [void]$libRows.Append("</tbody></table></div>")
                } elseif ($folder.Error) {
                    [void]$libRows.Append("<div class='sp-folder-detail' id='fp-$fRowIdx'><span class='sp-error'>$([System.Web.HttpUtility]::HtmlEncode($folder.Error))</span></div>")
                } else {
                    [void]$libRows.Append("<div class='sp-folder-detail' id='fp-$fRowIdx'><span class='sp-note'>Permissions inherited from library</span></div>")
                }
                [void]$libRows.Append("</div>")
            }
            [void]$libRows.Append("</div>")
        }
        [void]$libRows.Append("</div></div>")
    }

    # ── CSS ───────────────────────────────────────────────────
    $css = ":root{--bg:#1E1E2E;--bg2:#2D2D3F;--bg3:#1A1A2E;--bg4:#252535;--border:#4B5563;--border2:#374151;--text:#E2E8F0;--text2:#9CA3AF;--text3:#6B7280;--accent:#A78BFA;--accent2:#7C3AED;--green:#34D399;--red:#F87171;--yellow:#FBBF24;--blue:#60A5FA;--font-mono:Consolas,monospace;--font-ui:'Segoe UI',system-ui,sans-serif;--radius:6px}*{box-sizing:border-box;margin:0;padding:0}body{background:var(--bg);color:var(--text);font-family:var(--font-ui);font-size:13px;line-height:1.5}.header{background:linear-gradient(180deg,#2D2D3F,#1E1E2E);border-bottom:1px solid var(--border);padding:24px 32px 18px}.source-badge{display:inline-flex;align-items:center;gap:6px;background:rgba(96,165,250,.15);border:1px solid rgba(96,165,250,.3);color:var(--blue);font-size:11px;padding:3px 10px;border-radius:12px;margin-bottom:8px;font-weight:600}.title{font-size:20px;font-weight:700;margin-bottom:4px}.subtitle{color:var(--text2);font-family:var(--font-mono);font-size:11px;margin-bottom:14px;word-break:break-all}.meta{display:flex;flex-wrap:wrap;gap:18px}.meta-item{display:flex;flex-direction:column;gap:2px}.meta-lbl{color:var(--text3);font-size:10px;text-transform:uppercase;letter-spacing:.8px}.meta-val{font-size:13px;font-weight:500}.stats-bar{display:flex;gap:10px;padding:14px 32px;background:var(--bg2);border-bottom:1px solid var(--border);flex-wrap:wrap}.stat{background:var(--bg4);border:1px solid var(--border);border-radius:var(--radius);padding:9px 16px;display:flex;flex-direction:column;align-items:center;min-width:110px}.stat-num{font-size:22px;font-weight:700;line-height:1}.stat-lbl{font-size:11px;color:var(--text2);margin-top:2px}.s-site .stat-num{color:var(--accent)}.s-libs .stat-num{color:var(--blue)}.s-folders .stat-num{color:var(--green)}.s-ext .stat-num{color:var(--yellow)}.section{padding:20px 32px;border-bottom:1px solid var(--border2)}.section-title{font-size:13px;font-weight:700;color:var(--accent);text-transform:uppercase;letter-spacing:.6px;margin-bottom:12px}.sp-table{width:100%;border-collapse:collapse;font-size:12px}.sp-table thead tr{background:var(--bg4)}.sp-table th{text-align:left;padding:5px 10px;color:var(--text2);font-weight:500;font-size:11px;border-bottom:1px solid var(--border)}.sp-table td{padding:5px 10px;border-bottom:1px solid var(--border2);vertical-align:top}.sp-identity{font-family:var(--font-mono);font-size:11px;font-weight:500}.sp-type-badge{font-size:10px;padding:2px 7px;border-radius:10px;background:rgba(167,139,250,.15);border:1px solid rgba(167,139,250,.3);color:var(--accent)}.sp-roles{color:var(--text2)}.sp-count{color:var(--text2);text-align:center}.sp-members{color:var(--text2);font-size:11px}.ext-disabled{color:var(--green)}.ext-limited{color:var(--yellow)}.ext-open{color:var(--red)}.sp-library{border:1px solid var(--border2);border-radius:var(--radius);margin-bottom:6px;overflow:hidden}.sp-lib-header{display:flex;align-items:center;gap:8px;padding:8px 14px;cursor:pointer;background:var(--bg2);user-select:none}.sp-lib-header:hover{background:var(--bg4)}.sp-toggle{font-size:10px;color:var(--text3);transition:transform .2s}.sp-toggle-sm{font-size:9px;color:var(--text3);transition:transform .2s}.sp-lib-icon{font-size:15px}.sp-lib-name{font-family:var(--font-mono);font-size:12px;font-weight:600;flex:1}.sp-lib-type{font-size:10px;color:var(--text3);padding:2px 7px;background:var(--bg4);border-radius:10px}.sp-folder-count{font-size:10px;color:var(--text3)}.sp-lib-body{display:none;padding:12px 16px;border-top:1px solid var(--border2)}.sp-lib-body.open{display:block}.sp-section-title{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.6px;color:var(--text2);margin-bottom:6px}.sp-mini-table{width:100%;border-collapse:collapse;font-size:11px}.sp-mini-table th{text-align:left;padding:4px 8px;color:var(--text3);border-bottom:1px solid var(--border2);font-weight:500}.sp-mini-table td{padding:4px 8px;border-bottom:1px solid var(--border2)}.sp-folder-list{margin-top:4px}.sp-folder-row{border:1px solid var(--border2);border-radius:4px;margin-bottom:2px;overflow:hidden}.sp-folder-inherited{border-left:3px solid var(--border2)}.sp-folder-unique{border-left:3px solid var(--yellow);background:rgba(251,191,36,.04)}.sp-folder-error{border-left:3px solid var(--red)}.sp-folder-hdr{display:flex;align-items:center;gap:6px;padding:5px 10px;cursor:pointer;user-select:none}.sp-folder-hdr:hover{background:var(--bg2)}.sp-folder-icon{font-size:12px}.sp-folder-name{font-family:var(--font-mono);font-size:11px;flex:1}.sp-badge{display:inline-block;font-size:9px;padding:1px 6px;border-radius:8px;font-weight:500}.sp-badge-unique{background:rgba(251,191,36,.15);border:1px solid rgba(251,191,36,.3);color:var(--yellow)}.sp-badge-inherited{background:var(--bg4);border:1px solid var(--border2);color:var(--text3)}.sp-badge-error{background:rgba(241,113,113,.12);border:1px solid var(--red);color:var(--red)}.sp-badge-link{background:rgba(96,165,250,.12);border:1px solid rgba(96,165,250,.3);color:var(--blue)}.sp-folder-detail{display:none;padding:8px 12px;border-top:1px solid var(--border2);background:var(--bg3)}.sp-folder-detail.open{display:block}.sp-note{color:var(--text3);font-style:italic;font-size:11px}.sp-error{color:var(--red);font-size:11px}.libraries{padding:16px 32px}.footer{text-align:center;padding:16px;color:var(--text3);font-size:11px;border-top:1px solid var(--border2);background:var(--bg3)}"

    $js = "function toggleSP(id,hdr){var d=document.getElementById(id),ic=hdr.querySelector('.sp-toggle,.sp-toggle-sm');if(!d)return;if(d.classList.contains('open')){d.classList.remove('open');if(ic)ic.style.transform='';}else{d.classList.add('open');if(ic)ic.style.transform='rotate(90deg)';}}"

    $totalFolders = ($Drives | ForEach-Object { $_.Folders.Count } | Measure-Object -Sum).Sum

    $encodedSite = [System.Web.HttpUtility]::HtmlEncode($SiteUrl)
    $encodedName = [System.Web.HttpUtility]::HtmlEncode($SiteName)
    $encodedUser = [System.Web.HttpUtility]::HtmlEncode($currentUser)

    $html  = "<!DOCTYPE html><html lang='en'><head><meta charset='UTF-8'>"
    $html += "<title>ACLens SP &mdash; $encodedName</title><style>$css</style></head><body>"
    $html += "<div class='header'>"
    $html += "<div class='source-badge'>&#9729; SharePoint Online</div>"
    $html += "<div class='title'>ACLens &mdash; SharePoint Permission Report</div>"
    $html += "<div class='subtitle'>$encodedSite</div>"
    $html += "<div class='meta'>"
    $html += "<div class='meta-item'><span class='meta-lbl'>Site</span><span class='meta-val'>$encodedName</span></div>"
    $html += "<div class='meta-item'><span class='meta-lbl'>Created on</span><span class='meta-val'>$reportDate</span></div>"
    $html += "<div class='meta-item'><span class='meta-lbl'>Computer</span><span class='meta-val'>$computerName</span></div>"
    $html += "<div class='meta-item'><span class='meta-lbl'>User</span><span class='meta-val'>$encodedUser</span></div>"
    $html += "</div></div>"

    $html += "<div class='stats-bar'>"
    $html += "<div class='stat s-site'><span class='stat-num'>$($SitePerms.Count)</span><span class='stat-lbl'>Site permissions</span></div>"
    $html += "<div class='stat s-libs'><span class='stat-num'>$($Drives.Count)</span><span class='stat-lbl'>Libraries</span></div>"
    $html += "<div class='stat s-folders'><span class='stat-num'>$totalFolders</span><span class='stat-lbl'>Folders scanned</span></div>"
    $html += "<div class='stat s-ext'><span class='stat-num ext-badge $extClass'>$([System.Web.HttpUtility]::HtmlEncode($ExternalSharing))</span><span class='stat-lbl'>External sharing</span></div>"
    $html += "</div>"

    # Site Permissions
    $html += "<div class='section'><div class='section-title'>&#127968; Site-Level Permissions</div>"
    if ($SitePerms.Count -gt 0) {
        $html += "<table class='sp-table'><thead><tr><th>Principal</th><th>Type</th><th>Roles</th></tr></thead><tbody>"
        $html += $sitePermRows.ToString()
        $html += "</tbody></table>"
    } else {
        $html += "<span class='sp-note'>No site-level permissions found.</span>"
    }
    $html += "</div>"

    # Groups
    if ($Groups.Count -gt 0) {
        $html += "<div class='section'><div class='section-title'>&#128101; SharePoint Groups &amp; Members</div>"
        $html += "<table class='sp-table'><thead><tr><th>Group</th><th>Members</th><th>Details</th></tr></thead><tbody>"
        $html += $groupRows.ToString()
        $html += "</tbody></table></div>"
    }

    # Libraries
    $html += "<div class='libraries'>"
    $html += "<div class='section-title' style='margin-bottom:12px'>&#128218; Document Libraries &amp; Folder Permissions</div>"
    $html += $libRows.ToString()
    $html += "</div>"

    $html += "<div class='footer'>ACLens SharePoint Report &bull; $reportDate &bull; $computerName</div>"
    $html += "<script>$js</script></body></html>"

    [System.IO.File]::WriteAllText($OutputPath, $html, [System.Text.UTF8Encoding]::new($true))
    return [PSCustomObject]@{ Libraries = $Drives.Count; Folders = $totalFolders }
}


# ── SharePoint Compare HTML Report ───────────────────────────
function New-SPCompareHTMLReport {
    param([hashtable]$OldData, [hashtable]$NewData, [string]$OldLabel, [string]$NewLabel, [string]$OutputPath)

    $reportDate = Get-Date -Format "dd.MM.yyyy HH:mm:ss"
    $allPaths   = @($OldData.Keys) + @($NewData.Keys) | Sort-Object -Unique

    $added = 0; $removed = 0; $changed = 0; $same = 0
    $rows  = [System.Text.StringBuilder]::new()

    foreach ($path in $allPaths) {
        $inOld = $OldData.ContainsKey($path)
        $inNew = $NewData.ContainsKey($path)

        if (-not $inOld -and $inNew) {
            $added++
            [void]$rows.Append("<tr class='row-added'><td class='sp-diff-path'>+ $([System.Web.HttpUtility]::HtmlEncode($path))</td>")
            [void]$rows.Append("<td><span class='diff-badge badge-added'>Added</span></td>")
            [void]$rows.Append("<td class='diff-note'>New folder — not in baseline</td></tr>")
        }
        elseif ($inOld -and -not $inNew) {
            $removed++
            [void]$rows.Append("<tr class='row-removed'><td class='sp-diff-path'>- $([System.Web.HttpUtility]::HtmlEncode($path))</td>")
            [void]$rows.Append("<td><span class='diff-badge badge-removed'>Removed</span></td>")
            [void]$rows.Append("<td class='diff-note'>Folder removed — no longer exists</td></tr>")
        }
        else {
            $old = $OldData[$path]; $new = $NewData[$path]
            $oldPerms = ($old.Permissions | Sort-Object GrantedTo,Roles | ForEach-Object { "$($_.GrantedTo)|$($_.Roles)|$($_.IsInherited)" }) -join ";"
            $newPerms = ($new.Permissions | Sort-Object GrantedTo,Roles | ForEach-Object { "$($_.GrantedTo)|$($_.Roles)|$($_.IsInherited)" }) -join ";"
            $oldUniq  = $old.HasUniquePerms; $newUniq = $new.HasUniquePerms

            if ($oldPerms -ne $newPerms -or $oldUniq -ne $newUniq) {
                $changed++
                $details = [System.Text.StringBuilder]::new()

                if ($oldUniq -ne $newUniq) {
                    $oldU = if ($oldUniq) { "Unique" } else { "Inherited" }
                    $newU = if ($newUniq) { "Unique" } else { "Inherited" }
                    [void]$details.Append("<div class='sp-diff-detail'><span class='diff-key'>Permissions:</span> <span class='old-val'>$oldU</span> &rarr; <span class='new-val'>$newU</span></div>")
                }

                $oldSet = @{}; foreach ($p in $old.Permissions) { $oldSet["$($p.GrantedTo)|$($p.Roles)"] = $p }
                $newSet = @{}; foreach ($p in $new.Permissions) { $newSet["$($p.GrantedTo)|$($p.Roles)"] = $p }
                foreach ($k in ($newSet.Keys | Where-Object { -not $oldSet.ContainsKey($_) })) {
                    [void]$details.Append("<div class='sp-diff-detail added-rule'>+ $([System.Web.HttpUtility]::HtmlEncode($k))</div>")
                }
                foreach ($k in ($oldSet.Keys | Where-Object { -not $newSet.ContainsKey($_) })) {
                    [void]$details.Append("<div class='sp-diff-detail removed-rule'>- $([System.Web.HttpUtility]::HtmlEncode($k))</div>")
                }

                [void]$rows.Append("<tr class='row-changed'><td class='sp-diff-path'>~ $([System.Web.HttpUtility]::HtmlEncode($path))</td>")
                [void]$rows.Append("<td><span class='diff-badge badge-changed'>Changed</span></td>")
                [void]$rows.Append("<td>$($details.ToString())</td></tr>")
            } else {
                $same++
            }
        }
    }

    $css = ":root{--bg:#1E1E2E;--bg2:#2D2D3F;--bg3:#1A1A2E;--bg4:#252535;--border:#4B5563;--border2:#374151;--text:#E2E8F0;--text2:#9CA3AF;--text3:#6B7280;--accent:#A78BFA;--green:#34D399;--red:#F87171;--yellow:#FBBF24;--blue:#60A5FA;--font-mono:Consolas,monospace;--font-ui:'Segoe UI',system-ui,sans-serif;--radius:6px}*{box-sizing:border-box;margin:0;padding:0}body{background:var(--bg);color:var(--text);font-family:var(--font-ui);font-size:13px;line-height:1.5}.header{background:linear-gradient(180deg,#2D2D3F,#1E1E2E);border-bottom:1px solid var(--border);padding:24px 32px 18px}.source-badge{display:inline-flex;align-items:center;gap:6px;background:rgba(96,165,250,.15);border:1px solid rgba(96,165,250,.3);color:var(--blue);font-size:11px;padding:3px 10px;border-radius:12px;margin-bottom:8px;font-weight:600}.title{font-size:20px;font-weight:700;margin-bottom:6px}.meta{display:flex;flex-wrap:wrap;gap:18px;margin-top:10px}.meta-item{display:flex;flex-direction:column;gap:2px}.meta-lbl{color:var(--text3);font-size:10px;text-transform:uppercase;letter-spacing:.8px}.meta-val{font-size:13px;font-weight:500}.stats-bar{display:flex;gap:10px;padding:14px 32px;background:var(--bg2);border-bottom:1px solid var(--border);flex-wrap:wrap}.stat{background:var(--bg4);border:1px solid var(--border);border-radius:var(--radius);padding:9px 16px;display:flex;flex-direction:column;align-items:center;min-width:100px}.stat-num{font-size:22px;font-weight:700;line-height:1}.stat-lbl{font-size:11px;color:var(--text2);margin-top:2px}.s-changed .stat-num{color:var(--yellow)}.s-added .stat-num{color:var(--green)}.s-removed .stat-num{color:var(--red)}.s-same .stat-num{color:var(--text2)}.table-wrap{padding:16px 24px}.sp-diff-table{width:100%;border-collapse:collapse;font-size:12px}thead tr{background:var(--bg4)}th{text-align:left;padding:6px 10px;color:var(--text2);font-weight:500;font-size:11px;border-bottom:1px solid var(--border)}td{padding:6px 10px;border-bottom:1px solid var(--border2);vertical-align:top}.sp-diff-path{font-family:var(--font-mono);font-size:11px}.row-added{background:rgba(52,211,153,.05)}.row-added .sp-diff-path{color:var(--green)}.row-removed{background:rgba(241,113,113,.05)}.row-removed .sp-diff-path{color:var(--red)}.row-changed{background:rgba(251,191,36,.04)}.row-changed .sp-diff-path{color:var(--yellow)}.diff-badge{display:inline-block;font-size:10px;padding:2px 8px;border-radius:10px;font-weight:600}.badge-added{background:rgba(52,211,153,.15);border:1px solid rgba(52,211,153,.4);color:var(--green)}.badge-removed{background:rgba(241,113,113,.15);border:1px solid rgba(241,113,113,.4);color:var(--red)}.badge-changed{background:rgba(251,191,36,.15);border:1px solid rgba(251,191,36,.4);color:var(--yellow)}.diff-note{color:var(--text2)}.sp-diff-detail{font-size:11px;margin:2px 0;color:var(--text2)}.diff-key{font-weight:600;color:var(--text);margin-right:4px}.old-val{color:var(--red);text-decoration:line-through;margin-right:4px}.new-val{color:var(--green)}.added-rule{color:var(--green)}.removed-rule{color:var(--red)}.no-diff{padding:40px 32px;text-align:center;color:var(--text3);font-size:14px}.footer{text-align:center;padding:16px;color:var(--text3);font-size:11px;border-top:1px solid var(--border2);background:var(--bg3)}"

    $html  = "<!DOCTYPE html><html lang='en'><head><meta charset='UTF-8'><title>ACLens SP Compare</title><style>$css</style></head><body>"
    $html += "<div class='header'><div class='source-badge'>&#9729; SharePoint Online &mdash; Diff</div>"
    $html += "<div class='title'>ACLens &mdash; SharePoint Permission Diff</div>"
    $html += "<div class='meta'>"
    $html += "<div class='meta-item'><span class='meta-lbl'>Baseline</span><span class='meta-val'>$([System.Web.HttpUtility]::HtmlEncode($OldLabel))</span></div>"
    $html += "<div class='meta-item'><span class='meta-lbl'>Current Scan</span><span class='meta-val'>$([System.Web.HttpUtility]::HtmlEncode($NewLabel))</span></div>"
    $html += "<div class='meta-item'><span class='meta-lbl'>Generated</span><span class='meta-val'>$reportDate</span></div>"
    $html += "</div></div>"
    $html += "<div class='stats-bar'>"
    $html += "<div class='stat s-changed'><span class='stat-num'>$changed</span><span class='stat-lbl'>Permissions changed</span></div>"
    $html += "<div class='stat s-added'><span class='stat-num'>$added</span><span class='stat-lbl'>Folders added</span></div>"
    $html += "<div class='stat s-removed'><span class='stat-num'>$removed</span><span class='stat-lbl'>Folders removed</span></div>"
    $html += "<div class='stat s-same'><span class='stat-num'>$same</span><span class='stat-lbl'>Unchanged</span></div>"
    $html += "</div>"

    if (($changed + $added + $removed) -eq 0) {
        $html += "<div class='no-diff'>&#10003; No differences found — SharePoint permissions are identical.</div>"
    } else {
        $html += "<div class='table-wrap'><table class='sp-diff-table'>"
        $html += "<thead><tr><th>Path</th><th>Status</th><th>Details</th></tr></thead><tbody>"
        $html += $rows.ToString()
        $html += "</tbody></table></div>"
    }

    $html += "<div class='footer'>ACLens SharePoint Compare &bull; $reportDate</div></body></html>"
    [System.IO.File]::WriteAllText($OutputPath, $html, [System.Text.UTF8Encoding]::new($true))
}



