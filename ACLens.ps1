#Requires -Version 5.1
<#
.SYNOPSIS
    ACLens - NTFS Permission Analyzer & Reporter

.DESCRIPTION
    Entry point. Loads all modules and launches the GUI.

.NOTES
    Version:    0.2.1-alpha
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
    $aYes.Location = [System.Drawing.Point]::new(18, 214); $aYes.FlatStyle = "Flat"
    $aYes.Font = [System.Drawing.Font]::new("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $aYes.BackColor = HAC '#7C3AED'; $aYes.ForeColor = [System.Drawing.Color]::White
    $aYes.Cursor = "Hand"; $aYes.FlatAppearance.BorderColor = HAC '#4C1D95'
    $aYes.FlatAppearance.BorderSize = 1
    $aYes.DialogResult = [System.Windows.Forms.DialogResult]::Yes
    $adm.Controls.Add($aYes)

    $aNo = [System.Windows.Forms.Button]::new()
    $aNo.Text = "Continue without Admin"; $aNo.Size = [System.Drawing.Size]::new(180, 38)
    $aNo.Location = [System.Drawing.Point]::new(208, 214); $aNo.FlatStyle = "Flat"
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
$script:FN = [System.Drawing.Font]::new("Segoe UI",  9)
$script:FB = [System.Drawing.Font]::new("Segoe UI",  9,  [System.Drawing.FontStyle]::Bold)
$script:FT = [System.Drawing.Font]::new("Segoe UI", 15, [System.Drawing.FontStyle]::Bold)
$script:FS = [System.Drawing.Font]::new("Segoe UI",  8)

# ── Layout constants ─────────────────────────────────────────
$script:PAD   = 20
$script:HDR_H = 54
$script:FTR_H = 62
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
    $changedFolders = ($FolderData | Where-Object { $_.HasChanges   -eq $true }).Count
    $errorFolders   = ($FolderData | Where-Object { $null -ne $_.Error }).Count
    $inheritedOnly  = ($FolderData | Where-Object { $_.AllInherited -eq $true -and $_.HasChanges -eq $false }).Count

    # CSS and JS decoded from Base64 - avoids ALL PowerShell string escaping issues
    $cssB64 = (
        'OnJvb3R7LS1iZzojMUUxRTJFOy0tYmcyOiMyRDJEM0Y7LS1iZzM6IzFBMUEyRTstLWJnNDojMjUy' +

        'NTM1Oy0tYm9yZGVyOiM0QjU1NjM7LS1ib3JkZXIyOiMzNzQxNTE7LS10ZXh0OiNFMkU4RjA7LS10' +

        'ZXh0MjojOUNBM0FGOy0tdGV4dDM6IzZCNzI4MDstLWFjY2VudDojQTc4QkZBOy0tYWNjZW50Mjoj' +

        'N0MzQUVEOy0tYWNjZW50MzojNEMxRDk1Oy0tZ3JlZW46IzM0RDM5OTstLXJlZDojRjg3MTcxOy0t' +

        'eWVsbG93OiNGQkJGMjQ7LS1vcmFuZ2U6I0Y1OUUwQjstLWJsdWU6IzYwQTVGQTstLWZvbnQtbW9u' +

        'bzonQ2FzY2FkaWEgQ29kZScsJ0NvbnNvbGFzJyxtb25vc3BhY2U7LS1mb250LXVpOidTZWdvZSBV' +

        'SScsc3lzdGVtLXVpLHNhbnMtc2VyaWY7LS1yYWRpdXM6NnB4fQoqe2JveC1zaXppbmc6Ym9yZGVy' +

        'LWJveDttYXJnaW46MDtwYWRkaW5nOjB9CmJvZHl7YmFja2dyb3VuZDp2YXIoLS1iZyk7Y29sb3I6' +

        'dmFyKC0tdGV4dCk7Zm9udC1mYW1pbHk6dmFyKC0tZm9udC11aSk7Zm9udC1zaXplOjEzcHg7bGlu' +

        'ZS1oZWlnaHQ6MS41fQoucmVwb3J0LWhlYWRlcntiYWNrZ3JvdW5kOmxpbmVhci1ncmFkaWVudCgx' +

        'ODBkZWcsIzJEMkQzRiwjMUUxRTJFKTtib3JkZXItYm90dG9tOjFweCBzb2xpZCB2YXIoLS1ib3Jk' +

        'ZXIpO3BhZGRpbmc6MjRweCAzMnB4IDE4cHh9Ci5yZXBvcnQtdGl0bGV7Zm9udC1zaXplOjIwcHg7' +

        'Zm9udC13ZWlnaHQ6NzAwO2Rpc3BsYXk6ZmxleDthbGlnbi1pdGVtczpjZW50ZXI7Z2FwOjEwcHg7' +

        'bWFyZ2luLWJvdHRvbTo0cHh9Ci5yZXBvcnQtdGl0bGUgLmljb257Y29sb3I6dmFyKC0tYWNjZW50' +

        'KTtmb250LXNpemU6MjJweH0KLnJlcG9ydC1zdWJ0aXRsZXtjb2xvcjp2YXIoLS10ZXh0Mik7Zm9u' +

        'dC1mYW1pbHk6dmFyKC0tZm9udC1tb25vKTtmb250LXNpemU6MTFweDttYXJnaW4tYm90dG9tOjE0' +

        'cHg7d29yZC1icmVhazpicmVhay1hbGx9Ci5yZXBvcnQtbWV0YXtkaXNwbGF5OmZsZXg7ZmxleC13' +

        'cmFwOndyYXA7Z2FwOjE4cHh9Ci5tZXRhLWl0ZW17ZGlzcGxheTpmbGV4O2ZsZXgtZGlyZWN0aW9u' +

        'OmNvbHVtbjtnYXA6MnB4fQoubWV0YS1sYWJlbHtjb2xvcjp2YXIoLS10ZXh0Myk7Zm9udC1zaXpl' +

        'OjEwcHg7dGV4dC10cmFuc2Zvcm06dXBwZXJjYXNlO2xldHRlci1zcGFjaW5nOi44cHh9Ci5tZXRh' +

        'LXZhbHVle2NvbG9yOnZhcigtLXRleHQpO2ZvbnQtc2l6ZToxM3B4O2ZvbnQtd2VpZ2h0OjUwMH0K' +

        'LnN0YXRzLWJhcntkaXNwbGF5OmZsZXg7Z2FwOjEwcHg7cGFkZGluZzoxNHB4IDMycHg7YmFja2dy' +

        'b3VuZDp2YXIoLS1iZzIpO2JvcmRlci1ib3R0b206MXB4IHNvbGlkIHZhcigtLWJvcmRlcik7Zmxl' +

        'eC13cmFwOndyYXB9Ci5zdGF0LWNhcmR7YmFja2dyb3VuZDp2YXIoLS1iZzQpO2JvcmRlcjoxcHgg' +

        'c29saWQgdmFyKC0tYm9yZGVyKTtib3JkZXItcmFkaXVzOnZhcigtLXJhZGl1cyk7cGFkZGluZzo5' +

        'cHggMTZweDtkaXNwbGF5OmZsZXg7ZmxleC1kaXJlY3Rpb246Y29sdW1uO2FsaWduLWl0ZW1zOmNl' +

        'bnRlcjttaW4td2lkdGg6MTEwcHh9Ci5zdGF0LW51bWJlcntmb250LXNpemU6MjJweDtmb250LXdl' +

        'aWdodDo3MDA7bGluZS1oZWlnaHQ6MX0KLnN0YXQtbGFiZWx7Zm9udC1zaXplOjExcHg7Y29sb3I6' +

        'dmFyKC0tdGV4dDIpO21hcmdpbi10b3A6MnB4fQouc3RhdC10b3RhbCAuc3RhdC1udW1iZXJ7Y29s' +

        'b3I6dmFyKC0tYWNjZW50KX0uc3RhdC1jaGFuZ2VkIC5zdGF0LW51bWJlcntjb2xvcjp2YXIoLS15' +

        'ZWxsb3cpfS5zdGF0LWluaGVyaXRlZCAuc3RhdC1udW1iZXJ7Y29sb3I6dmFyKC0tZ3JlZW4pfS5z' +

        'dGF0LWVycm9yIC5zdGF0LW51bWJlcntjb2xvcjp2YXIoLS1yZWQpfQouYWdlbmRhLWJhcntiYWNr' +

        'Z3JvdW5kOnZhcigtLWJnMyk7Ym9yZGVyLWJvdHRvbToxcHggc29saWQgdmFyKC0tYm9yZGVyKX0K' +

        'LmFnZW5kYS10b2dnbGV7d2lkdGg6MTAwJTtkaXNwbGF5OmZsZXg7YWxpZ24taXRlbXM6Y2VudGVy' +

        'O2dhcDo4cHg7cGFkZGluZzoxMHB4IDMycHg7YmFja2dyb3VuZDpub25lO2JvcmRlcjpub25lO2Nv' +

        'bG9yOnZhcigtLXRleHQyKTtmb250LXNpemU6MTJweDtmb250LWZhbWlseTp2YXIoLS1mb250LXVp' +

        'KTtjdXJzb3I6cG9pbnRlcjt0ZXh0LWFsaWduOmxlZnR9Ci5hZ2VuZGEtdG9nZ2xlOmhvdmVye2Jh' +

        'Y2tncm91bmQ6cmdiYSgxNjcsMTM5LDI1MCwuMDgpO2NvbG9yOnZhcigtLWFjY2VudCl9Ci5hZ2Vu' +

        'ZGEtdG9nZ2xlLWljb257Zm9udC1zaXplOjlweDt0cmFuc2l0aW9uOnRyYW5zZm9ybSAuMnN9Ci5h' +

        'Z2VuZGEtdG9nZ2xlLWhpbnR7bWFyZ2luLWxlZnQ6YXV0bztmb250LXNpemU6MTBweDtjb2xvcjp2' +

        'YXIoLS10ZXh0Myl9Ci5hZ2VuZGEtY29udGVudHtwYWRkaW5nOjE2cHggMzJweCAyMHB4O2JvcmRl' +

        'ci10b3A6MXB4IHNvbGlkIHZhcigtLWJvcmRlcjIpO2JhY2tncm91bmQ6dmFyKC0tYmcpfQouYWdl' +

        'bmRhLWdyaWR7ZGlzcGxheTpncmlkO2dyaWQtdGVtcGxhdGUtY29sdW1uczpyZXBlYXQoYXV0by1m' +

        'aWxsLG1pbm1heCgzMDBweCwxZnIpKTtnYXA6MTJweH0KLmFnZW5kYS1zZWN0aW9ue2JhY2tncm91' +

        'bmQ6dmFyKC0tYmc0KTtib3JkZXI6MXB4IHNvbGlkIHZhcigtLWJvcmRlcjIpO2JvcmRlci1yYWRp' +

        'dXM6dmFyKC0tcmFkaXVzKTtwYWRkaW5nOjEycHggMTRweH0KLmFnZW5kYS1zZWN0aW9uLXRpdGxl' +

        'e2ZvbnQtc2l6ZToxMHB4O2ZvbnQtd2VpZ2h0OjcwMDt0ZXh0LXRyYW5zZm9ybTp1cHBlcmNhc2U7' +

        'bGV0dGVyLXNwYWNpbmc6LjhweDtjb2xvcjp2YXIoLS1hY2NlbnQpO21hcmdpbi1ib3R0b206OHB4' +

        'O3BhZGRpbmctYm90dG9tOjVweDtib3JkZXItYm90dG9tOjFweCBzb2xpZCB2YXIoLS1ib3JkZXIy' +

        'KX0KLmFnZW5kYS1yb3d7ZGlzcGxheTpmbGV4O2FsaWduLWl0ZW1zOmJhc2VsaW5lO2dhcDo3cHg7' +

        'bWFyZ2luLWJvdHRvbTo1cHg7Zm9udC1zaXplOjExcHg7bGluZS1oZWlnaHQ6MS40fQouYWdlbmRh' +

        'LXJvdzpsYXN0LWNoaWxke21hcmdpbi1ib3R0b206MH0KLmFnZW5kYS1kb3R7d2lkdGg6OHB4O2hl' +

        'aWdodDo4cHg7Ym9yZGVyLXJhZGl1czoycHg7ZmxleC1zaHJpbms6MDttYXJnaW4tdG9wOjNweDtk' +

        'aXNwbGF5OmlubGluZS1ibG9ja30KLmFnZW5kYS1sYWJlbHtjb2xvcjp2YXIoLS10ZXh0KTtmb250' +

        'LXdlaWdodDo2MDA7d2hpdGUtc3BhY2U6bm93cmFwO2ZsZXgtc2hyaW5rOjB9Ci5hZ2VuZGEtcGVy' +

        'bXtmb250LWZhbWlseTp2YXIoLS1mb250LW1vbm8pO2ZvbnQtc2l6ZToxMHB4O2NvbG9yOnZhcigt' +

        'LXRleHQyKTtmb250LXdlaWdodDo1MDA7d2hpdGUtc3BhY2U6bm93cmFwO2ZsZXgtc2hyaW5rOjB9' +

        'Ci5hZ2VuZGEtZGVzY3tjb2xvcjp2YXIoLS10ZXh0Mik7Zm9udC1zaXplOjExcHh9Ci5hZ2VuZGEt' +

        'YmFkZ2V7ZmxleC1zaHJpbms6MH0KLnRvb2xiYXJ7ZGlzcGxheTpmbGV4O2FsaWduLWl0ZW1zOmNl' +

        'bnRlcjtnYXA6OHB4O3BhZGRpbmc6MTBweCAzMnB4O2JhY2tncm91bmQ6dmFyKC0tYmczKTtib3Jk' +

        'ZXItYm90dG9tOjFweCBzb2xpZCB2YXIoLS1ib3JkZXIyKTtmbGV4LXdyYXA6d3JhcH0KLnRvb2xi' +

        'YXItbGFiZWx7Y29sb3I6dmFyKC0tdGV4dDIpO2ZvbnQtc2l6ZToxMnB4fQouYnRue2JhY2tncm91' +

        'bmQ6dmFyKC0tYmcyKTtib3JkZXI6MXB4IHNvbGlkIHZhcigtLWJvcmRlcik7Ym9yZGVyLXJhZGl1' +

        'czp2YXIoLS1yYWRpdXMpO2NvbG9yOnZhcigtLXRleHQpO2N1cnNvcjpwb2ludGVyO2ZvbnQtc2l6' +

        'ZToxMnB4O3BhZGRpbmc6NXB4IDEycHg7dHJhbnNpdGlvbjpiYWNrZ3JvdW5kIC4xNXMsYm9yZGVy' +

        'LWNvbG9yIC4xNXM7Zm9udC1mYW1pbHk6dmFyKC0tZm9udC11aSl9Ci5idG46aG92ZXJ7YmFja2dy' +

        'b3VuZDp2YXIoLS1iZzQpO2JvcmRlci1jb2xvcjp2YXIoLS1hY2NlbnQpfQouYnRuLWFjY2VudHti' +

        'YWNrZ3JvdW5kOnZhcigtLWFjY2VudDIpO2JvcmRlci1jb2xvcjp2YXIoLS1hY2NlbnQzKTtjb2xv' +

        'cjp2YXIoLS10ZXh0KX0KLmJ0bi1hY2NlbnQ6aG92ZXJ7YmFja2dyb3VuZDp2YXIoLS1hY2NlbnQz' +

        'KX0KLmZpbHRlci1ncm91cHtkaXNwbGF5OmZsZXg7Z2FwOjVweH0KLmZpbHRlci1idG57cGFkZGlu' +

        'Zzo0cHggMTBweDtmb250LXNpemU6MTFweDtib3JkZXItcmFkaXVzOjIwcHh9Ci5maWx0ZXItYnRu' +

        'LmFjdGl2ZXtiYWNrZ3JvdW5kOnZhcigtLWFjY2VudDIpO2JvcmRlci1jb2xvcjp2YXIoLS1hY2Nl' +

        'bnQzKTtjb2xvcjp2YXIoLS10ZXh0KX0KLnNlYXJjaC1ib3h7YmFja2dyb3VuZDp2YXIoLS1iZzIp' +

        'O2JvcmRlcjoxcHggc29saWQgdmFyKC0tYm9yZGVyKTtib3JkZXItcmFkaXVzOnZhcigtLXJhZGl1' +

        'cyk7Y29sb3I6dmFyKC0tdGV4dCk7Zm9udC1zaXplOjEycHg7cGFkZGluZzo1cHggMTBweDt3aWR0' +

        'aDoyNDBweDtvdXRsaW5lOm5vbmU7Zm9udC1mYW1pbHk6dmFyKC0tZm9udC11aSl9Ci5zZWFyY2gt' +

        'Ym94OmZvY3Vze2JvcmRlci1jb2xvcjp2YXIoLS1hY2NlbnQpfQouc2VhcmNoLWJveDo6cGxhY2Vo' +

        'b2xkZXJ7Y29sb3I6dmFyKC0tdGV4dDMpfQouZm9sZGVyLWxpc3R7cGFkZGluZzoxMHB4IDE4cHh9' +

        'Ci5mb2xkZXItcm93e2JvcmRlcjoxcHggc29saWQgdmFyKC0tYm9yZGVyMik7Ym9yZGVyLXJhZGl1' +

        'czp2YXIoLS1yYWRpdXMpO21hcmdpbi1ib3R0b206M3B4O292ZXJmbG93OmhpZGRlbn0KLmZvbGRl' +

        'ci1yb3cuc3RhdHVzLXJvb3R7Ym9yZGVyLWxlZnQ6M3B4IHNvbGlkIHZhcigtLWFjY2VudCk7YmFj' +

        'a2dyb3VuZDpyZ2JhKDE2NywxMzksMjUwLC4wNSl9Ci5mb2xkZXItcm93LnN0YXR1cy1jaGFuZ2Vk' +

        'e2JvcmRlci1sZWZ0OjNweCBzb2xpZCB2YXIoLS15ZWxsb3cpO2JhY2tncm91bmQ6cmdiYSgyNTEs' +

        'MTkxLDM2LC4wNSl9Ci5mb2xkZXItcm93LnN0YXR1cy1lcnJvcntib3JkZXItbGVmdDozcHggc29s' +

        'aWQgdmFyKC0tcmVkKX0KLmZvbGRlci1yb3cuc3RhdHVzLWluaGVyaXRlZHtib3JkZXItbGVmdDoz' +

        'cHggc29saWQgdmFyKC0tYm9yZGVyMil9Ci5mb2xkZXItaGVhZGVye2Rpc3BsYXk6ZmxleDthbGln' +

        'bi1pdGVtczpjZW50ZXI7anVzdGlmeS1jb250ZW50OnNwYWNlLWJldHdlZW47cGFkZGluZzo3cHgg' +

        'MTJweDtjdXJzb3I6cG9pbnRlcjt1c2VyLXNlbGVjdDpub25lO2dhcDo4cHh9Ci5mb2xkZXItaGVh' +

        'ZGVyOmhvdmVye2JhY2tncm91bmQ6cmdiYSgxNjcsMTM5LDI1MCwuMDYpfQouZm9sZGVyLWxlZnR7' +

        'ZGlzcGxheTpmbGV4O2FsaWduLWl0ZW1zOmNlbnRlcjtnYXA6N3B4O2ZsZXg6MTttaW4td2lkdGg6' +

        'MH0KLmZvbGRlci1yaWdodHtkaXNwbGF5OmZsZXg7YWxpZ24taXRlbXM6Y2VudGVyO2dhcDo3cHg7' +

        'ZmxleC1zaHJpbms6MH0KLnRvZ2dsZS1pY29ue2NvbG9yOnZhcigtLXRleHQzKTtmb250LXNpemU6' +

        'MTBweDt0cmFuc2l0aW9uOnRyYW5zZm9ybSAuMnM7ZmxleC1zaHJpbms6MH0KLmZvbGRlci1pY29u' +

        'e2ZvbnQtc2l6ZToxNHB4O2ZsZXgtc2hyaW5rOjB9Ci5mb2xkZXItcGF0aHtmb250LWZhbWlseTp2' +

        'YXIoLS1mb250LW1vbm8pO2ZvbnQtc2l6ZToxMnB4O2NvbG9yOnZhcigtLXRleHQpO3doaXRlLXNw' +

        'YWNlOm5vd3JhcDtvdmVyZmxvdzpoaWRkZW47dGV4dC1vdmVyZmxvdzplbGxpcHNpc30KLnN0YXR1' +

        'cy1iYWRnZSwuaW5oZXJpdC1iYWRnZSwub3duZXItYmFkZ2V7Zm9udC1zaXplOjEwcHg7cGFkZGlu' +

        'ZzoycHggOHB4O2JvcmRlci1yYWRpdXM6MTJweDt3aGl0ZS1zcGFjZTpub3dyYXA7Zm9udC13ZWln' +

        'aHQ6NTAwfQoub3duZXItYmFkZ2V7YmFja2dyb3VuZDp2YXIoLS1iZzQpO2JvcmRlcjoxcHggc29s' +

        'aWQgdmFyKC0tYm9yZGVyMik7Y29sb3I6dmFyKC0tdGV4dDIpfQouc3RhdHVzLWJhZGdlLnN0YXR1' +

        'cy1yb290e2JhY2tncm91bmQ6cmdiYSgxNjcsMTM5LDI1MCwuMTUpO2JvcmRlcjoxcHggc29saWQg' +

        'dmFyKC0tYWNjZW50Mik7Y29sb3I6dmFyKC0tYWNjZW50KX0KLnN0YXR1cy1iYWRnZS5zdGF0dXMt' +

        'Y2hhbmdlZHtiYWNrZ3JvdW5kOnJnYmEoMjUxLDE5MSwzNiwuMTUpO2JvcmRlcjoxcHggc29saWQg' +

        'dmFyKC0teWVsbG93KTtjb2xvcjp2YXIoLS15ZWxsb3cpfQouc3RhdHVzLWJhZGdlLnN0YXR1cy1l' +

        'cnJvcntiYWNrZ3JvdW5kOnJnYmEoMjQxLDExMywxMTMsLjEyKTtib3JkZXI6MXB4IHNvbGlkIHZh' +

        'cigtLXJlZCk7Y29sb3I6dmFyKC0tcmVkKX0KLnN0YXR1cy1iYWRnZS5zdGF0dXMtaW5oZXJpdGVk' +

        'e2JhY2tncm91bmQ6cmdiYSg1MiwyMTEsMTUzLC4xMik7Ym9yZGVyOjFweCBzb2xpZCByZ2JhKDUy' +

        'LDIxMSwxNTMsLjMpO2NvbG9yOnZhcigtLWdyZWVuKX0KLmluaGVyaXQteWVze2JhY2tncm91bmQ6' +

        'cmdiYSg1MiwyMTEsMTUzLC4xMik7Ym9yZGVyOjFweCBzb2xpZCByZ2JhKDUyLDIxMSwxNTMsLjMp' +

        'O2NvbG9yOnZhcigtLWdyZWVuKX0KLmluaGVyaXQtbm97YmFja2dyb3VuZDpyZ2JhKDI0MSwxMTMs' +

        'MTEzLC4xMik7Ym9yZGVyOjFweCBzb2xpZCByZ2JhKDI0MSwxMTMsMTEzLC4zKTtjb2xvcjp2YXIo' +

        'LS1yZWQpfQouZm9sZGVyLWRldGFpbHN7ZGlzcGxheTpub25lO3BhZGRpbmc6MTBweCAxOHB4IDE0' +

        'cHg7Ym9yZGVyLXRvcDoxcHggc29saWQgdmFyKC0tYm9yZGVyMik7YmFja2dyb3VuZDp2YXIoLS1i' +

        'ZzMpfQouZm9sZGVyLWRldGFpbHMub3BlbntkaXNwbGF5OmJsb2NrfQouZGV0YWlscy1tZXRhe2Rp' +

        'c3BsYXk6ZmxleDtmbGV4LXdyYXA6d3JhcDtnYXA6MTRweDttYXJnaW4tYm90dG9tOjEycHg7cGFk' +

        'ZGluZzo5cHggMTJweDtiYWNrZ3JvdW5kOnZhcigtLWJnNCk7Ym9yZGVyLXJhZGl1czp2YXIoLS1y' +

        'YWRpdXMpO2JvcmRlcjoxcHggc29saWQgdmFyKC0tYm9yZGVyMik7Zm9udC1zaXplOjEycHg7Y29s' +

        'b3I6dmFyKC0tdGV4dDIpfQouZGV0YWlscy1tZXRhIHN0cm9uZ3tjb2xvcjp2YXIoLS10ZXh0KX0K' +

        'LnNhbWUtYXMtcGFyZW50e2NvbG9yOnZhcigtLXRleHQzKTtmb250LXN0eWxlOml0YWxpYztmb250' +

        'LXNpemU6MTJweDtwYWRkaW5nOjdweCAxMnB4O2JhY2tncm91bmQ6dmFyKC0tYmc0KTtib3JkZXIt' +

        'cmFkaXVzOnZhcigtLXJhZGl1cyk7Ym9yZGVyOjFweCBkYXNoZWQgdmFyKC0tYm9yZGVyMil9Ci5l' +

        'cnJvci1ib3h7YmFja2dyb3VuZDpyZ2JhKDI0MSwxMTMsMTEzLC4wNyk7Ym9yZGVyOjFweCBzb2xp' +

        'ZCB2YXIoLS1yZWQpO2JvcmRlci1yYWRpdXM6dmFyKC0tcmFkaXVzKTtjb2xvcjp2YXIoLS1yZWQp' +

        'O3BhZGRpbmc6OXB4IDEycHg7Zm9udC1zaXplOjEycHh9Ci5lcnJvci1ib3ggY29kZXtmb250LWZh' +

        'bWlseTp2YXIoLS1mb250LW1vbm8pO2ZvbnQtc2l6ZToxMXB4fQoucnVsZXMtc2VjdGlvbnttYXJn' +

        'aW4tYm90dG9tOjEwcHh9Ci5ydWxlcy1zZWN0aW9uLXRpdGxle2ZvbnQtc2l6ZToxMXB4O2ZvbnQt' +

        'd2VpZ2h0OjYwMDt0ZXh0LXRyYW5zZm9ybTp1cHBlcmNhc2U7bGV0dGVyLXNwYWNpbmc6LjZweDtw' +

        'YWRkaW5nOjNweCA5cHg7bWFyZ2luLWJvdHRvbTo1cHg7Ym9yZGVyLXJhZGl1czozcHg7ZGlzcGxh' +

        'eTppbmxpbmUtYmxvY2t9Ci5leHBsaWNpdC10aXRsZXtjb2xvcjp2YXIoLS1hY2NlbnQpO2JhY2tn' +

        'cm91bmQ6cmdiYSgxNjcsMTM5LDI1MCwuMTIpfQouaW5oZXJpdGVkLXRpdGxle2NvbG9yOnZhcigt' +

        'LXRleHQzKTtiYWNrZ3JvdW5kOnZhcigtLWJnNCl9Ci5ydWxlcy10YWJsZXt3aWR0aDoxMDAlO2Jv' +

        'cmRlci1jb2xsYXBzZTpjb2xsYXBzZTtmb250LXNpemU6MTJweH0KLnJ1bGVzLXRhYmxlIHRoZWFk' +

        'IHRye2JhY2tncm91bmQ6dmFyKC0tYmc0KX0KLnJ1bGVzLXRhYmxlIHRoe3RleHQtYWxpZ246bGVm' +

        'dDtwYWRkaW5nOjVweCA5cHg7Y29sb3I6dmFyKC0tdGV4dDIpO2ZvbnQtd2VpZ2h0OjUwMDtmb250' +

        'LXNpemU6MTFweDtib3JkZXItYm90dG9tOjFweCBzb2xpZCB2YXIoLS1ib3JkZXIpfQoucnVsZXMt' +

        'dGFibGUgdGR7cGFkZGluZzo1cHggOXB4O2JvcmRlci1ib3R0b206MXB4IHNvbGlkIHZhcigtLWJv' +

        'cmRlcjIpO3ZlcnRpY2FsLWFsaWduOm1pZGRsZX0KLnJ1bGVzLXRhYmxlIHRib2R5IHRyOmhvdmVy' +

        'e2JhY2tncm91bmQ6cmdiYSgxNjcsMTM5LDI1MCwuMDQpfQoucnVsZXMtdGFibGUgdGJvZHkgdHI6' +

        'bGFzdC1jaGlsZCB0ZHtib3JkZXItYm90dG9tOm5vbmV9Ci5pbmhlcml0ZWQtdGFibGV7b3BhY2l0' +

        'eTouNzJ9Ci5kZW55LXJ1bGV7YmFja2dyb3VuZDpyZ2JhKDI0MSwxMTMsMTEzLC4wNikhaW1wb3J0' +

        'YW50fQouZGVueS1ydWxlIHRke2NvbG9yOnZhcigtLXJlZCkhaW1wb3J0YW50fQouaWRlbnRpdHl7' +

        'Zm9udC1mYW1pbHk6dmFyKC0tZm9udC1tb25vKTtmb250LXNpemU6MTFweDtmb250LXdlaWdodDo1' +

        'MDB9Ci5yaWdodHN7Y29sb3I6dmFyKC0tdGV4dCl9LnNjb3Ble2NvbG9yOnZhcigtLXRleHQyKTtm' +

        'b250LXNpemU6MTFweH0KLm5vLXJ1bGVze2NvbG9yOnZhcigtLXRleHQzKTtmb250LXN0eWxlOml0' +

        'YWxpYztmb250LXNpemU6MTJweDtwYWRkaW5nOjVweCA5cHh9Ci5iYWRnZXtkaXNwbGF5OmlubGlu' +

        'ZS1ibG9jaztmb250LXNpemU6MTBweDtwYWRkaW5nOjJweCA3cHg7Ym9yZGVyLXJhZGl1czoxMHB4' +

        'O2ZvbnQtd2VpZ2h0OjUwMDt3aGl0ZS1zcGFjZTpub3dyYXB9Ci5iYWRnZS1hbGxvd3tiYWNrZ3Jv' +

        'dW5kOnJnYmEoNTIsMjExLDE1MywuMTIpO2JvcmRlcjoxcHggc29saWQgcmdiYSg1MiwyMTEsMTUz' +

        'LC4zNSk7Y29sb3I6dmFyKC0tZ3JlZW4pfQouYmFkZ2UtZGVueXtiYWNrZ3JvdW5kOnJnYmEoMjQx' +

        'LDExMywxMTMsLjEyKTtib3JkZXI6MXB4IHNvbGlkIHJnYmEoMjQxLDExMywxMTMsLjM1KTtjb2xv' +

        'cjp2YXIoLS1yZWQpfQouYmFkZ2UtaW5oZXJpdGVke2JhY2tncm91bmQ6dmFyKC0tYmcyKTtib3Jk' +

        'ZXI6MXB4IHNvbGlkIHZhcigtLWJvcmRlcjIpO2NvbG9yOnZhcigtLXRleHQzKX0KLmJhZGdlLWV4' +

        'cGxpY2l0e2JhY2tncm91bmQ6cmdiYSgxNjcsMTM5LDI1MCwuMTUpO2JvcmRlcjoxcHggc29saWQg' +

        'cmdiYSgxNjcsMTM5LDI1MCwuMyk7Y29sb3I6dmFyKC0tYWNjZW50KX0KLnJlcG9ydC1mb290ZXJ7' +

        'dGV4dC1hbGlnbjpjZW50ZXI7cGFkZGluZzoxNnB4O2NvbG9yOnZhcigtLXRleHQzKTtmb250LXNp' +

        'emU6MTFweDtib3JkZXItdG9wOjFweCBzb2xpZCB2YXIoLS1ib3JkZXIyKTtiYWNrZ3JvdW5kOnZh' +

        'cigtLWJnMyl9CkBtZWRpYSBwcmludHsudG9vbGJhciwuc3RhdHMtYmFye2Rpc3BsYXk6bm9uZX0u' +

        'Zm9sZGVyLWRldGFpbHN7ZGlzcGxheTpibG9jayFpbXBvcnRhbnR9LmZvbGRlci1yb3d7YnJlYWst' +

        'aW5zaWRlOmF2b2lkfX0='
    )
    $jsB64 = (
        'ZnVuY3Rpb24gdG9nZ2xlQWdlbmRhKGJ0bil7dmFyIGM9ZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQo' +

        'J2FnZW5kYUNvbnRlbnQnKSxpYz1idG4ucXVlcnlTZWxlY3RvcignLmFnZW5kYS10b2dnbGUtaWNv' +

        'bicpLGhpbnQ9YnRuLnF1ZXJ5U2VsZWN0b3IoJy5hZ2VuZGEtdG9nZ2xlLWhpbnQnKSxoaWRkZW49' +

        'Yy5zdHlsZS5kaXNwbGF5PT09J25vbmUnfHxjLnN0eWxlLmRpc3BsYXk9PT0nJztjLnN0eWxlLmRp' +

        'c3BsYXk9aGlkZGVuPydibG9jayc6J25vbmUnO2ljLnN0eWxlLnRyYW5zZm9ybT1oaWRkZW4/J3Jv' +

        'dGF0ZSg5MGRlZyknOicnO2hpbnQudGV4dENvbnRlbnQ9aGlkZGVuPydjbGljayB0byBjb2xsYXBz' +

        'ZSc6J2NsaWNrIHRvIGV4cGFuZCc7fQpmdW5jdGlvbiB0b2dnbGUoaWQsaGRyKXt2YXIgZD1kb2N1' +

        'bWVudC5nZXRFbGVtZW50QnlJZChpZCksaWM9aGRyLnF1ZXJ5U2VsZWN0b3IoJy50b2dnbGUtaWNv' +

        'bicpO2lmKGQuY2xhc3NMaXN0LmNvbnRhaW5zKCdvcGVuJykpe2QuY2xhc3NMaXN0LnJlbW92ZSgn' +

        'b3BlbicpO2ljLnN0eWxlLnRyYW5zZm9ybT0nJzt9ZWxzZXtkLmNsYXNzTGlzdC5hZGQoJ29wZW4n' +

        'KTtpYy5zdHlsZS50cmFuc2Zvcm09J3JvdGF0ZSg5MGRlZyknO319CmZ1bmN0aW9uIGV4cGFuZEFs' +

        'bCgpe2RvY3VtZW50LnF1ZXJ5U2VsZWN0b3JBbGwoJy5mb2xkZXItZGV0YWlscycpLmZvckVhY2go' +

        'ZnVuY3Rpb24oZCl7ZC5jbGFzc0xpc3QuYWRkKCdvcGVuJyk7fSk7ZG9jdW1lbnQucXVlcnlTZWxl' +

        'Y3RvckFsbCgnLnRvZ2dsZS1pY29uJykuZm9yRWFjaChmdW5jdGlvbihpKXtpLnN0eWxlLnRyYW5z' +

        'Zm9ybT0ncm90YXRlKDkwZGVnKSc7fSl9CmZ1bmN0aW9uIGNvbGxhcHNlQWxsKCl7ZG9jdW1lbnQu' +

        'cXVlcnlTZWxlY3RvckFsbCgnLmZvbGRlci1kZXRhaWxzJykuZm9yRWFjaChmdW5jdGlvbihkKXtk' +

        'LmNsYXNzTGlzdC5yZW1vdmUoJ29wZW4nKTt9KTtkb2N1bWVudC5xdWVyeVNlbGVjdG9yQWxsKCcu' +

        'dG9nZ2xlLWljb24nKS5mb3JFYWNoKGZ1bmN0aW9uKGkpe2kuc3R5bGUudHJhbnNmb3JtPScnO30p' +

        'fQp2YXIgX2Y9J2FsbCcsX3M9Jyc7CmZ1bmN0aW9uIGZpbHRlclJvd3MoZixidG4pe19mPWY7ZG9j' +

        'dW1lbnQucXVlcnlTZWxlY3RvckFsbCgnLmZpbHRlci1idG4nKS5mb3JFYWNoKGZ1bmN0aW9uKGIp' +

        'e2IuY2xhc3NMaXN0LnJlbW92ZSgnYWN0aXZlJyk7fSk7YnRuLmNsYXNzTGlzdC5hZGQoJ2FjdGl2' +

        'ZScpO2FwcGx5VmlzKCk7fQpmdW5jdGlvbiBzZWFyY2hSb3dzKHYpe19zPXYudG9Mb3dlckNhc2Uo' +

        'KTthcHBseVZpcygpO30KZnVuY3Rpb24gYXBwbHlWaXMoKXtkb2N1bWVudC5xdWVyeVNlbGVjdG9y' +

        'QWxsKCcuZm9sZGVyLXJvdycpLmZvckVhY2goZnVuY3Rpb24ocm93KXt2YXIgc2hvdz10cnVlO2lm' +

        'KF9mPT09J2NoYW5nZWQnJiYhcm93LmNsYXNzTGlzdC5jb250YWlucygnc3RhdHVzLWNoYW5nZWQn' +

        'KSlzaG93PWZhbHNlO2lmKF9mPT09J2luaGVyaXRlZCcmJiFyb3cuY2xhc3NMaXN0LmNvbnRhaW5z' +

        'KCdzdGF0dXMtaW5oZXJpdGVkJykpc2hvdz1mYWxzZTtpZihfZj09PSdlcnJvcicmJiFyb3cuY2xh' +

        'c3NMaXN0LmNvbnRhaW5zKCdzdGF0dXMtZXJyb3InKSlzaG93PWZhbHNlO2lmKHNob3cmJl9zJiZy' +

        'b3cudGV4dENvbnRlbnQudG9Mb3dlckNhc2UoKS5pbmRleE9mKF9zKTwwKXNob3c9ZmFsc2U7cm93' +

        'LnN0eWxlLmRpc3BsYXk9c2hvdz8nJzonbm9uZSc7fSk7fQpkb2N1bWVudC5xdWVyeVNlbGVjdG9y' +

        'QWxsKCcuZm9sZGVyLXJvdy5zdGF0dXMtY2hhbmdlZCwuZm9sZGVyLXJvdy5zdGF0dXMtcm9vdCcp' +

        'LmZvckVhY2goZnVuY3Rpb24ocm93KXt2YXIgZD1yb3cucXVlcnlTZWxlY3RvcignLmZvbGRlci1k' +

        'ZXRhaWxzJyksaWM9cm93LnF1ZXJ5U2VsZWN0b3IoJy50b2dnbGUtaWNvbicpO2lmKGQpZC5jbGFz' +

        'c0xpc3QuYWRkKCdvcGVuJyk7aWYoaWMpaWMuc3R5bGUudHJhbnNmb3JtPSdyb3RhdGUoOTBkZWcp' +

        'Jzt9KTs='
    )
    $css = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($cssB64))
    $js  = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($jsB64))

    # ── Build folder rows ────────────────────────────────────
    $rows = [System.Text.StringBuilder]::new()
    $rowIndex = 0
    foreach ($folder in $FolderData) {
        $rowIndex++
        $depth       = $folder.Depth
        $isRoot      = ($depth -eq 0)
        $nbsp        = [string]::new([char]0x00A0, $depth * 4)
        $displayPath = if ($isRoot) { $folder.Path } else { $folder.Path.Replace($RootPath, '') }

        if ($null -ne $folder.Error) {
            $sClass = 'status-error';     $sIcon = '&#9888;';   $sTxt = 'Error'
        } elseif ($isRoot) {
            $sClass = 'status-root';      $sIcon = '&#127968;'; $sTxt = 'Root'
        } elseif ($folder.HasChanges) {
            $sClass = 'status-changed';   $sIcon = '&#9888;';   $sTxt = 'Changed'
        } else {
            $sClass = 'status-inherited'; $sIcon = '&#8595;';   $sTxt = 'Inherited'
        }

        $iIcon  = if ($folder.InheritanceEnabled) { '&#10003;' } else { '&#10007;' }
        $iClass = if ($folder.InheritanceEnabled) { 'inherit-yes' } else { 'inherit-no' }
        $iTxt   = if ($folder.InheritanceEnabled) { 'Inheritance active' } else { 'Inheritance disabled' }
        $prot   = if ($folder.AreAccessRulesProtected) { 'Yes (inheritance broken)' } else { 'No' }
        $ownerE = [System.Web.HttpUtility]::HtmlEncode($folder.Owner)
        $pathE  = [System.Web.HttpUtility]::HtmlEncode($displayPath)
        $fullE  = [System.Web.HttpUtility]::HtmlEncode($folder.Path)
        $ownerH = if ($folder.Owner) { "<span class='owner-badge'>&#128100; $ownerE</span>" } else { '' }
        $openCl = if ($isRoot -or $folder.HasChanges) { ' open' } else { '' }
        $iconRt = if ($isRoot -or $folder.HasChanges) { " style='transform:rotate(90deg)'" } else { '' }

        [void]$rows.Append("<div class='folder-row depth-$depth $sClass'>")
        [void]$rows.Append("<div class='folder-header' onclick='toggle(`"details-$rowIndex`",this)'>")
        [void]$rows.Append("<div class='folder-left'><span class='toggle-icon'$iconRt>&#9658;</span>")
        [void]$rows.Append("<span class='folder-icon'>&#128193;</span>")
        [void]$rows.Append("<span class='folder-path'>$nbsp$pathE</span></div>")
        [void]$rows.Append("<div class='folder-right'>$ownerH")
        [void]$rows.Append("<span class='inherit-badge $iClass'>$iIcon $iTxt</span>")
        [void]$rows.Append("<span class='status-badge $sClass'>$sIcon $sTxt</span>")
        [void]$rows.Append("</div></div>")
        [void]$rows.Append("<div class='folder-details$openCl' id='details-$rowIndex'>")
        [void]$rows.Append("<div class='details-meta'>")
        [void]$rows.Append("<span><strong>Path:</strong> $fullE</span>")
        [void]$rows.Append("<span><strong>Owner:</strong> $ownerE</span>")
        [void]$rows.Append("<span><strong>Inheritance:</strong> $iTxt</span>")
        [void]$rows.Append("<span><strong>Rules protected:</strong> $prot</span>")
        [void]$rows.Append("</div>")

        if ($null -ne $folder.Error) {
            $errE = [System.Web.HttpUtility]::HtmlEncode($folder.Error)
            [void]$rows.Append("<div class='error-box'>&#9888; Error: <code>$errE</code></div>")
        } elseif ($isRoot -or $folder.HasChanges) {
            $explicit  = @($folder.Rules | Where-Object { -not $_.IsInherited })
            $inherited = @($folder.Rules | Where-Object { $_.IsInherited })
            if ($explicit.Count -gt 0 -or (-not $folder.InheritanceEnabled)) {
                [void]$rows.Append("<div class='rules-section'><div class='rules-section-title explicit-title'>&#128274; Explicit Permissions</div>")
                if ($explicit.Count -gt 0) {
                    [void]$rows.Append("<table class='rules-table'><thead><tr><th>Principal</th><th>Type</th><th>Permission</th><th>Applies to</th><th>Kind</th></tr></thead><tbody>")
                    foreach ($r in $explicit) {
                        $idE = [System.Web.HttpUtility]::HtmlEncode($r.Identity)
                        $dc  = if ($r.AccessType -eq 'Deny') { " class='deny-rule'" } else { '' }
                        $tl  = if ($r.AccessType -eq 'Deny') { "<span class='badge badge-deny'>Deny</span>" } else { "<span class='badge badge-allow'>Allow</span>" }
                        [void]$rows.Append("<tr$dc><td class='identity'>$idE</td><td>$tl</td><td class='rights'>$($r.RightsSimple)</td><td class='scope'>$($r.InheritanceDesc)</td><td><span class='badge badge-explicit'>Explicit</span></td></tr>")
                    }
                    [void]$rows.Append("</tbody></table>")
                } else {
                    [void]$rows.Append("<div class='no-rules'>No explicit permissions</div>")
                }
                [void]$rows.Append("</div>")
            }
            if ($inherited.Count -gt 0) {
                [void]$rows.Append("<div class='rules-section'><div class='rules-section-title inherited-title'>&#128275; Inherited Permissions</div>")
                [void]$rows.Append("<table class='rules-table inherited-table'><thead><tr><th>Principal</th><th>Type</th><th>Permission</th><th>Applies to</th><th>Kind</th></tr></thead><tbody>")
                foreach ($r in $inherited) {
                    $idE = [System.Web.HttpUtility]::HtmlEncode($r.Identity)
                    $dc  = if ($r.AccessType -eq 'Deny') { " class='deny-rule'" } else { '' }
                    $tl  = if ($r.AccessType -eq 'Deny') { "<span class='badge badge-deny'>Deny</span>" } else { "<span class='badge badge-allow'>Allow</span>" }
                    [void]$rows.Append("<tr$dc><td class='identity'>$idE</td><td>$tl</td><td class='rights'>$($r.RightsSimple)</td><td class='scope'>$($r.InheritanceDesc)</td><td><span class='badge badge-inherited'>Inherited</span></td></tr>")
                }
                [void]$rows.Append("</tbody></table></div>")
            }
        } else {
            [void]$rows.Append("<div class='same-as-parent'>Permissions identical to parent folder &#8593;</div>")
        }
        [void]$rows.Append("</div></div>`n")
    }

    # ── Agenda ───────────────────────────────────────────────
    $agenda  = "<div class='agenda-bar'>"
    $agenda += "<button class='agenda-toggle' onclick='toggleAgenda(this)'>"
    $agenda += "<span class='agenda-toggle-icon'>&#9658;</span> &#128218; Legend &amp; Reference "
    $agenda += "<span class='agenda-toggle-hint'>click to expand</span></button>"
    $agenda += "<div class='agenda-content' id='agendaContent' style='display:none'><div class='agenda-grid'>"
    $agenda += "<div class='agenda-section'><div class='agenda-section-title'>Folder Status</div>"
    $agenda += "<div class='agenda-row'><span class='agenda-dot' style='background:var(--accent)'></span><span class='agenda-label'>Root</span><span class='agenda-desc'>The start folder of the analysis.</span></div>"
    $agenda += "<div class='agenda-row'><span class='agenda-dot' style='background:var(--yellow)'></span><span class='agenda-label'>Changed</span><span class='agenda-desc'>Permissions differ from parent &mdash; rules or inheritance modified.</span></div>"
    $agenda += "<div class='agenda-row'><span class='agenda-dot' style='background:var(--green)'></span><span class='agenda-label'>Inherited only</span><span class='agenda-desc'>All permissions inherited. No local changes.</span></div>"
    $agenda += "<div class='agenda-row'><span class='agenda-dot' style='background:var(--red)'></span><span class='agenda-label'>Error</span><span class='agenda-desc'>Could not read permissions &mdash; access denied.</span></div></div>"
    $agenda += "<div class='agenda-section'><div class='agenda-section-title'>Permission Types</div>"
    $agenda += "<div class='agenda-row'><span class='badge badge-allow agenda-badge'>Allow</span><span class='agenda-desc'>Grants the right to the principal.</span></div>"
    $agenda += "<div class='agenda-row'><span class='badge badge-deny agenda-badge'>Deny</span><span class='agenda-desc'>Blocks the right. Deny always overrides Allow.</span></div>"
    $agenda += "<div class='agenda-row'><span class='badge badge-explicit agenda-badge'>Explicit</span><span class='agenda-desc'>Set directly on this folder.</span></div>"
    $agenda += "<div class='agenda-row'><span class='badge badge-inherited agenda-badge'>Inherited</span><span class='agenda-desc'>Passed down from a parent folder.</span></div></div>"
    $agenda += "<div class='agenda-section'><div class='agenda-section-title'>Permission Levels</div>"
    $agenda += "<div class='agenda-row'><span class='agenda-perm'>Full Control</span><span class='agenda-desc'>Read, write, delete, change permissions, take ownership.</span></div>"
    $agenda += "<div class='agenda-row'><span class='agenda-perm'>Modify</span><span class='agenda-desc'>Read, write, delete. Cannot change permissions.</span></div>"
    $agenda += "<div class='agenda-row'><span class='agenda-perm'>Read &amp; Execute</span><span class='agenda-desc'>View and run files.</span></div>"
    $agenda += "<div class='agenda-row'><span class='agenda-perm'>Read</span><span class='agenda-desc'>View folder contents only.</span></div>"
    $agenda += "<div class='agenda-row'><span class='agenda-perm'>Write</span><span class='agenda-desc'>Create files and subfolders.</span></div>"
    $agenda += "<div class='agenda-row'><span class='agenda-perm'>Special</span><span class='agenda-desc'>Custom combination of rights.</span></div></div>"
    $agenda += "<div class='agenda-section'><div class='agenda-section-title'>Scope</div>"
    $agenda += "<div class='agenda-row'><span class='agenda-perm'>This folder, subfolders and files</span><span class='agenda-desc'>Applies everywhere below.</span></div>"
    $agenda += "<div class='agenda-row'><span class='agenda-perm'>This folder and subfolders</span><span class='agenda-desc'>Not files directly.</span></div>"
    $agenda += "<div class='agenda-row'><span class='agenda-perm'>This folder only</span><span class='agenda-desc'>Only this specific folder.</span></div>"
    $agenda += "<div class='agenda-row'><span class='agenda-perm'>Subfolders and files</span><span class='agenda-desc'>Contents only, not the folder itself.</span></div></div>"
    $agenda += "<div class='agenda-section'><div class='agenda-section-title'>Inheritance</div>"
    $agenda += "<div class='agenda-row'><span class='inherit-badge inherit-yes agenda-badge'>&#10003; Active</span><span class='agenda-desc'>Inherits permissions from parent.</span></div>"
    $agenda += "<div class='agenda-row'><span class='inherit-badge inherit-no agenda-badge'>&#10007; Disabled</span><span class='agenda-desc'>Managed independently &mdash; inheritance broken.</span></div>"
    $agenda += "<div class='agenda-row'><span class='agenda-perm'>Rules protected: Yes</span><span class='agenda-desc'>Inherited rules converted to explicit.</span></div></div>"
    $agenda += "</div></div></div>"

    # ── Toolbar ──────────────────────────────────────────────
    $toolbar  = "<div class='toolbar'><span class='toolbar-label'>Filter:</span><div class='filter-group'>"
    $toolbar += "<button class='btn filter-btn active' onclick='filterRows(`"all`",this)'>All</button>"
    $toolbar += "<button class='btn filter-btn' onclick='filterRows(`"changed`",this)' style='border-color:var(--yellow);color:var(--yellow)'>Changed</button>"
    $toolbar += "<button class='btn filter-btn' onclick='filterRows(`"inherited`",this)' style='border-color:rgba(52,211,153,.3);color:var(--green)'>Inherited only</button>"
    $toolbar += "<button class='btn filter-btn' onclick='filterRows(`"error`",this)' style='border-color:var(--red);color:var(--red)'>Errors</button>"
    $toolbar += "</div><input class='search-box' type='text' placeholder='Search path or user...' oninput='searchRows(this.value)'>"
    $toolbar += "<button class='btn btn-accent' onclick='expandAll()'>Expand all</button>"
    $toolbar += "<button class='btn' onclick='collapseAll()'>Collapse all</button>"
    $toolbar += "<button class='btn' onclick='window.print()'>Print</button></div>"

    # ── Assemble HTML ────────────────────────────────────────
    $encodedRoot = [System.Web.HttpUtility]::HtmlEncode($RootPath)
    $encodedUser = [System.Web.HttpUtility]::HtmlEncode($currentUser)

    $out = [System.Text.StringBuilder]::new()
    [void]$out.Append("<!DOCTYPE html><html lang='en'><head><meta charset='UTF-8'>")
    [void]$out.Append("<meta name='viewport' content='width=device-width,initial-scale=1.0'>")
    [void]$out.Append("<title>ACLens - $encodedRoot</title>")
    [void]$out.Append("<style>$css</style></head><body>")
    [void]$out.Append("<div class='report-header'>")
    [void]$out.Append("<div class='report-title'><span class='icon'>&#128274;</span>ACLens - NTFS Permission Report</div>")
    [void]$out.Append("<div class='report-subtitle'>$encodedRoot</div>")
    [void]$out.Append("<div class='report-meta'>")
    [void]$out.Append("<div class='meta-item'><span class='meta-label'>Created on</span><span class='meta-value'>$reportDate</span></div>")
    [void]$out.Append("<div class='meta-item'><span class='meta-label'>Computer</span><span class='meta-value'>$computerName</span></div>")
    [void]$out.Append("<div class='meta-item'><span class='meta-label'>User</span><span class='meta-value'>$encodedUser</span></div>")
    [void]$out.Append("<div class='meta-item'><span class='meta-label'>Total folders</span><span class='meta-value'>$totalFolders</span></div>")
    [void]$out.Append("</div></div>")
    [void]$out.Append("<div class='stats-bar'>")
    [void]$out.Append("<div class='stat-card stat-total'><span class='stat-number'>$totalFolders</span><span class='stat-label'>Total folders</span></div>")
    [void]$out.Append("<div class='stat-card stat-changed'><span class='stat-number'>$changedFolders</span><span class='stat-label'>Permissions changed</span></div>")
    [void]$out.Append("<div class='stat-card stat-inherited'><span class='stat-number'>$inheritedOnly</span><span class='stat-label'>Inherited only</span></div>")
    [void]$out.Append("<div class='stat-card stat-error'><span class='stat-number'>$errorFolders</span><span class='stat-label'>Errors</span></div>")
    [void]$out.Append("</div>")
    [void]$out.Append($agenda)
    [void]$out.Append($toolbar)
    [void]$out.Append("<div class='folder-list'>")
    [void]$out.Append($rows.ToString())
    [void]$out.Append("</div>")
    [void]$out.Append("<div class='report-footer'>ACLens &bull; $reportDate &bull; $computerName</div>")
    [void]$out.Append("<script>$js</script>")
    [void]$out.Append("</body></html>")

    [System.IO.File]::WriteAllText($OutputPath, $out.ToString(), [System.Text.UTF8Encoding]::new($true))
}

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
$C_CANCEL   = HC '#991B1B'
$C_CANC2    = HC '#7F1D1D'
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
$form.MinimumSize     = [System.Drawing.Size]::new(700, 440)
$form.StartPosition   = "CenterScreen"
$form.FormBorderStyle = "Sizable"
$form.MaximizeBox     = $true
$form.BackColor       = $C_BG
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
$lblVersion.Text      = "v0.2.1-alpha"
$lblVersion.Font      = [System.Drawing.Font]::new("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$lblVersion.ForeColor = HC '#A78BFA'
$lblVersion.BackColor = [System.Drawing.Color]::Transparent
$lblVersion.AutoSize  = $true
$lblVersion.Location  = [System.Drawing.Point]::new(100, 14)   # will be adjusted after lblTitle loads
$pnlHdr.Controls.Add($lblVersion)

# Subtitle — right of version badge
$lblSub = [System.Windows.Forms.Label]::new()
$lblSub.Text      = "NTFS & SharePoint Permission Analyzer"
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
$pnlFtr.Anchor    = "Bottom,Left,Right"
$form.Controls.Add($pnlFtr)

$sepFtr = [System.Windows.Forms.Panel]::new()
$sepFtr.BackColor = $C_BORDER
$sepFtr.Anchor    = "Bottom,Left,Right"
$form.Controls.Add($sepFtr)

$btnStart = [System.Windows.Forms.Button]::new()
$btnStart.Text      = "Start Analysis"
$btnStart.Size      = [System.Drawing.Size]::new(148, 36)
$btnStart.Location  = [System.Drawing.Point]::new(16, 13)
$btnStart.FlatStyle = "Flat"
$btnStart.Font      = $FB
$btnStart.BackColor = $C_PRIMARY
$btnStart.ForeColor = $C_TXT
$btnStart.Cursor    = "Hand"
$btnStart.FlatAppearance.BorderColor = $C_PRIM2
$btnStart.FlatAppearance.BorderSize  = 1
$pnlFtr.Controls.Add($btnStart)

$btnCancel = [System.Windows.Forms.Button]::new()
$btnCancel.Text      = "Cancel"
$btnCancel.Size      = [System.Drawing.Size]::new(88, 36)
$btnCancel.Location  = [System.Drawing.Point]::new(172, 13)
$btnCancel.FlatStyle = "Flat"
$btnCancel.Font      = $FB
$btnCancel.BackColor = $C_CANCEL
$btnCancel.ForeColor = [System.Drawing.Color]::White
$btnCancel.Cursor    = "Hand"
$btnCancel.Enabled   = $false
$btnCancel.FlatAppearance.BorderColor = $C_CANC2
$btnCancel.FlatAppearance.BorderSize  = 1
$pnlFtr.Controls.Add($btnCancel)

$btnOpenLast = [System.Windows.Forms.Button]::new()
$btnOpenLast.Text      = "Open Last Report"
$btnOpenLast.Size      = [System.Drawing.Size]::new(130, 36)
$btnOpenLast.Location  = [System.Drawing.Point]::new(270, 13)
$btnOpenLast.FlatStyle = "Flat"
$btnOpenLast.Font      = $FN
$btnOpenLast.BackColor = HC '#1D4ED8'
$btnOpenLast.ForeColor = [System.Drawing.Color]::White
$btnOpenLast.Cursor    = "Hand"
$btnOpenLast.Enabled   = $false
$btnOpenLast.FlatAppearance.BorderColor = HC '#1E40AF'
$btnOpenLast.FlatAppearance.BorderSize  = 1
$pnlFtr.Controls.Add($btnOpenLast)

$btnChangelog = [System.Windows.Forms.Button]::new()
$btnChangelog.Text      = "Changelog"
$btnChangelog.Size      = [System.Drawing.Size]::new(100, 36)
$btnChangelog.Location  = [System.Drawing.Point]::new(418, 13)
$btnChangelog.FlatStyle = "Flat"
$btnChangelog.Font      = $FN
$btnChangelog.BackColor = HC '#7C3AED'
$btnChangelog.ForeColor = [System.Drawing.Color]::White
$btnChangelog.Cursor    = "Hand"
$btnChangelog.FlatAppearance.BorderColor = HC '#4C1D95'
$btnChangelog.FlatAppearance.BorderSize  = 1
$pnlFtr.Controls.Add($btnChangelog)

$btnCompare = [System.Windows.Forms.Button]::new()
$btnCompare.Text      = "Compare Scans"
$btnCompare.Size      = [System.Drawing.Size]::new(120, 36)
$btnCompare.Location  = [System.Drawing.Point]::new(526, 13)
$btnCompare.FlatStyle = "Flat"
$btnCompare.Font      = $FN
$btnCompare.BackColor = HC '#1D4ED8'
$btnCompare.ForeColor = [System.Drawing.Color]::White
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
$tabNTFS.FlatStyle = "Flat"; $tabNTFS.Font = $FB; $tabNTFS.Cursor = "Hand"
$tabNTFS.BackColor = HC '#1E1E2E'; $tabNTFS.ForeColor = HC '#A78BFA'; $tabNTFS.FlatAppearance.BorderSize = 0
$tabNTFS.FlatAppearance.BorderSize = 0
$pnlTabs.Controls.Add($tabNTFS)

$tabSP = [System.Windows.Forms.Button]::new()
$tabSP.Text = "  SharePoint"; $tabSP.Size = [System.Drawing.Size]::new(140, 36)
$tabSP.FlatStyle = "Flat"; $tabSP.Font = $FB; $tabSP.Cursor = "Hand"
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
    if ($tab -eq "NTFS") {
        $tabNTFS.BackColor = HC '#1E1E2E'; $tabNTFS.ForeColor = HC '#A78BFA'; $tabNTFS.FlatAppearance.BorderSize = 0
        $tabSP.BackColor   = HC '#1A1A2E'; $tabSP.ForeColor   = HC '#6B7280'
        $pnlIn.Visible = $true;  $pnlSP.Visible = $false
    } else {
        $tabSP.BackColor   = HC '#1E1E2E'; $tabSP.ForeColor   = HC '#A78BFA'
        $tabNTFS.BackColor = HC '#1A1A2E'; $tabNTFS.ForeColor = HC '#6B7280'
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
    $b.FlatStyle = "Flat"
    $b.Font      = $FN
    $b.BackColor = $C_BTN
    $b.ForeColor = $C_TXT
    $b.Cursor    = "Hand"
    $b.FlatAppearance.BorderColor = $C_BORD2
    $b.FlatAppearance.BorderSize  = 1
    return $b
}

# Zeile 1: Start Path
$lblPath   = New-Label "Start Path:" $true
$tbPath    = New-TB $null
$btnBrowse = New-Btn "Browse..."
$pnlIn.Controls.Add($lblPath)
$pnlIn.Controls.Add($tbPath)
$pnlIn.Controls.Add($btnBrowse)

# Zeile 2: Output File
$lblOut  = New-Label "Output File (HTML + JSON):" $true
$tbOut   = New-TB "(automatic — saved to script folder)"
$btnSave = New-Btn "Save as..."
$pnlIn.Controls.Add($lblOut)
$pnlIn.Controls.Add($tbOut)
$pnlIn.Controls.Add($btnSave)

# Zeile 3: Depth
$lblDepth = New-Label "Maximum Depth:" $true
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
$chk.Text      = "Open report in browser after creation"
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
$lblStatus.Text      = "Ready."
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
$btnCancel.add_Click({ $script:cancelFlag = $true })

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
        $btnCancel.Enabled = $true
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
            $btnCancel.Enabled = $false
        }
        return
    }

    $rootPath = $tbPath.Text.Trim()
    if ([string]::IsNullOrEmpty($rootPath) -or -not (Test-Path $rootPath -PathType Container)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please enter a valid folder path.", "Invalid Path", "OK", "Warning") | Out-Null
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
    $btnCancel.Enabled = $true
    $script:cancelFlag = $false
    $progBar.Value     = 0
    $lblStats.Text     = ""
    $lblStatus.ForeColor = $C_MID

    try {
        $lblStatus.Text = "Step 1/3: Scanning folders..."
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
                $lblStatus.Text      = "Cancelled."
                $lblStatus.ForeColor = $C_WARN
                $progBar.Value       = 0
                $btnStart.Enabled    = $true
                $btnCancel.Enabled   = $false
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

        $lblStatus.Text = "Step 3/3: Generating HTML report..."
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
        $btnCancel.Enabled = $false
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
$spLblUrl.Text = "SharePoint Site URL:"; $spLblUrl.Font = $FB
$spLblUrl.ForeColor = HC '#E2E8F0'; $spLblUrl.BackColor = [System.Drawing.Color]::Transparent; $spLblUrl.AutoSize = $true
$pnlSP.Controls.Add($spLblUrl)

$spTbUrl = [System.Windows.Forms.TextBox]::new()
$spTbUrl.BackColor = HC '#2D2D3F'; $spTbUrl.ForeColor = HC '#6B7280'
$spTbUrl.Text = "https://yourtenant.sharepoint.com/sites/yoursite"
$spTbUrl.BorderStyle = "FixedSingle"; $spTbUrl.Font = $FN; $spTbUrl.Height = 26
$pnlSP.Controls.Add($spTbUrl)

# SP: Depth
$spLblDepth = [System.Windows.Forms.Label]::new()
$spLblDepth.Text = "Maximum Depth:"; $spLblDepth.Font = $FB
$spLblDepth.ForeColor = HC '#E2E8F0'; $spLblDepth.BackColor = [System.Drawing.Color]::Transparent; $spLblDepth.AutoSize = $true
$pnlSP.Controls.Add($spLblDepth)

$spNumDepth = [System.Windows.Forms.NumericUpDown]::new()
$spNumDepth.Size = [System.Drawing.Size]::new(68, 26); $spNumDepth.Minimum = 0; $spNumDepth.Maximum = 100; $spNumDepth.Value = 0
$spNumDepth.BackColor = HC '#2D2D3F'; $spNumDepth.ForeColor = HC '#E2E8F0'; $spNumDepth.Font = $FN
$pnlSP.Controls.Add($spNumDepth)

$spLblUnlim = [System.Windows.Forms.Label]::new()
$spLblUnlim.Text = "  0 = unlimited"; $spLblUnlim.Font = $FS
$spLblUnlim.ForeColor = HC '#6B7280'; $spLblUnlim.BackColor = [System.Drawing.Color]::Transparent; $spLblUnlim.AutoSize = $true
$pnlSP.Controls.Add($spLblUnlim)

# SP: Checkbox
$spChk = [System.Windows.Forms.CheckBox]::new()
$spChk.Text = "Open report in browser after creation"; $spChk.Font = $FN
$spChk.ForeColor = HC '#9CA3AF'; $spChk.BackColor = [System.Drawing.Color]::Transparent
$spChk.Checked = $true; $spChk.AutoSize = $true
$pnlSP.Controls.Add($spChk)

# SP: Setup button
$btnSPSetup = [System.Windows.Forms.Button]::new()
$btnSPSetup.Text = "Setup / Reconnect"; $btnSPSetup.Size = [System.Drawing.Size]::new(160, 26)
$btnSPSetup.FlatStyle = "Flat"; $btnSPSetup.Font = $FN
$btnSPSetup.BackColor = HC '#7C3AED'; $btnSPSetup.ForeColor = [System.Drawing.Color]::White; $btnSPSetup.Cursor = "Hand"
$btnSPSetup.FlatAppearance.BorderColor = HC '#4C1D95'; $btnSPSetup.FlatAppearance.BorderSize = 1
$pnlSP.Controls.Add($btnSPSetup)

function Update-SPStatus {
    $creds = Get-SPCredentials
    if ($creds.IsConfigured) {
        $spLblConn.Text      = "Connected"
        $spLblConn.ForeColor = HC '#34D399'
        $spLblConnDetail.Text = "Tenant: $($creds.TenantId)  |  Client: $($creds.ClientId)"
    } else {
        $spLblConn.Text      = "Not configured"
        $spLblConn.ForeColor = HC '#F87171'
        $spLblConnDetail.Text = "Click 'Setup / Reconnect' to register ACLens with your Microsoft 365 tenant."
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
    $wiz.Text            = "ACLens — SharePoint Setup"
    $wiz.ClientSize      = [System.Drawing.Size]::new(620, 480)
    $wiz.StartPosition   = "CenterParent"
    $wiz.FormBorderStyle = "FixedDialog"
    $wiz.MaximizeBox     = $false
    $wiz.BackColor       = HC '#1E1E2E'
    $wiz.Font            = $FN

    # Header
    $wHdr = [System.Windows.Forms.Panel]::new()
    $wHdr.SetBounds(0, 0, 620, 56); $wHdr.BackColor = HC '#1A1A2E'
    $wiz.Controls.Add($wHdr)

    $wTitle = [System.Windows.Forms.Label]::new()
    $wTitle.Text = "SharePoint Setup"; $wTitle.Font = [System.Drawing.Font]::new("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $wTitle.ForeColor = HC '#A78BFA'; $wTitle.BackColor = [System.Drawing.Color]::Transparent
    $wTitle.AutoSize = $true; $wTitle.Location = [System.Drawing.Point]::new(16, 13)
    $wHdr.Controls.Add($wTitle)

    $wSub = [System.Windows.Forms.Label]::new()
    $wSub.Text = "Connect ACLens to Microsoft 365"
    $wSub.Font = $FS; $wSub.ForeColor = HC '#6B7280'
    $wSub.BackColor = [System.Drawing.Color]::Transparent; $wSub.AutoSize = $true
    $wSub.Location = [System.Drawing.Point]::new(190, 21)
    $wHdr.Controls.Add($wSub)

    $wSepHdr = [System.Windows.Forms.Panel]::new()
    $wSepHdr.SetBounds(0, 56, 620, 1); $wSepHdr.BackColor = HC '#4B5563'
    $wiz.Controls.Add($wSepHdr)

    # Tab buttons: Auto (Device Code) | Manual
    $wTabAuto = [System.Windows.Forms.Button]::new()
    $wTabAuto.Text = "  ⚡  Automatic Setup"; $wTabAuto.Size = [System.Drawing.Size]::new(200, 34)
    $wTabAuto.Location = [System.Drawing.Point]::new(16, 68); $wTabAuto.FlatStyle = "Flat"
    $wTabAuto.Font = $FB; $wTabAuto.BackColor = HC '#7C3AED'; $wTabAuto.ForeColor = [System.Drawing.Color]::White
    $wTabAuto.Cursor = "Hand"; $wTabAuto.FlatAppearance.BorderColor = HC '#4C1D95'; $wTabAuto.FlatAppearance.BorderSize = 1
    $wiz.Controls.Add($wTabAuto)

    $wTabManual = [System.Windows.Forms.Button]::new()
    $wTabManual.Text = "  📋  Manual Setup"; $wTabManual.Size = [System.Drawing.Size]::new(180, 34)
    $wTabManual.Location = [System.Drawing.Point]::new(224, 68); $wTabManual.FlatStyle = "Flat"
    $wTabManual.Font = $FB; $wTabManual.BackColor = HC '#374151'; $wTabManual.ForeColor = HC '#9CA3AF'
    $wTabManual.Cursor = "Hand"; $wTabManual.FlatAppearance.BorderColor = HC '#4B5563'; $wTabManual.FlatAppearance.BorderSize = 1
    $wiz.Controls.Add($wTabManual)

    $wSepTab = [System.Windows.Forms.Panel]::new()
    $wSepTab.SetBounds(0, 103, 620, 1); $wSepTab.BackColor = HC '#374151'
    $wiz.Controls.Add($wSepTab)

    # ── Auto panel ───────────────────────────────────────────
    $pAuto = [System.Windows.Forms.Panel]::new()
    $pAuto.SetBounds(0, 104, 620, 316); $pAuto.BackColor = HC '#1E1E2E'
    $wiz.Controls.Add($pAuto)

    $aInfo = [System.Windows.Forms.Label]::new()
    $aInfo.Text = "ACLens will automatically create an App Registration in your Azure AD tenant using the Device Code flow. You need an account with Application Administrator or Global Administrator role."
    $aInfo.Font = $FN; $aInfo.ForeColor = HC '#9CA3AF'; $aInfo.BackColor = [System.Drawing.Color]::Transparent
    $aInfo.Size = [System.Drawing.Size]::new(580, 52); $aInfo.Location = [System.Drawing.Point]::new(20, 14)
    $pAuto.Controls.Add($aInfo)

    $aStep1 = [System.Windows.Forms.Label]::new()
    $aStep1.Text = "Step 1 — Click 'Get Device Code' below"
    $aStep1.Font = $FB; $aStep1.ForeColor = HC '#E2E8F0'; $aStep1.BackColor = [System.Drawing.Color]::Transparent
    $aStep1.AutoSize = $true; $aStep1.Location = [System.Drawing.Point]::new(20, 74)
    $pAuto.Controls.Add($aStep1)

    $aStep2 = [System.Windows.Forms.Label]::new()
    $aStep2.Text = "Step 2 — Open the link and enter the code shown below"
    $aStep2.Font = $FB; $aStep2.ForeColor = HC '#E2E8F0'; $aStep2.BackColor = [System.Drawing.Color]::Transparent
    $aStep2.AutoSize = $true; $aStep2.Location = [System.Drawing.Point]::new(20, 98)
    $pAuto.Controls.Add($aStep2)

    $aStep3 = [System.Windows.Forms.Label]::new()
    $aStep3.Text = "Step 3 — Sign in with your admin account — ACLens handles the rest"
    $aStep3.Font = $FB; $aStep3.ForeColor = HC '#E2E8F0'; $aStep3.BackColor = [System.Drawing.Color]::Transparent
    $aStep3.AutoSize = $true; $aStep3.Location = [System.Drawing.Point]::new(20, 122)
    $pAuto.Controls.Add($aStep3)

    # Device code box
    $aCodeBox = [System.Windows.Forms.Panel]::new()
    $aCodeBox.SetBounds(20, 152, 580, 64); $aCodeBox.BackColor = HC '#1A1A2E'
    $pAuto.Controls.Add($aCodeBox)

    $aCodeLabel = [System.Windows.Forms.Label]::new()
    $aCodeLabel.Text = "Click 'Get Device Code' to start"
    $aCodeLabel.Font = [System.Drawing.Font]::new("Consolas", 14, [System.Drawing.FontStyle]::Bold)
    $aCodeLabel.ForeColor = HC '#FBBF24'; $aCodeLabel.BackColor = [System.Drawing.Color]::Transparent
    $aCodeLabel.AutoSize = $true; $aCodeLabel.Location = [System.Drawing.Point]::new(16, 18)
    $aCodeBox.Controls.Add($aCodeLabel)

    $aUrlLabel = [System.Windows.Forms.Label]::new()
    $aUrlLabel.Text = ""; $aUrlLabel.Font = $FS
    $aUrlLabel.ForeColor = HC '#60A5FA'; $aUrlLabel.BackColor = [System.Drawing.Color]::Transparent
    $aUrlLabel.AutoSize = $true; $aUrlLabel.Location = [System.Drawing.Point]::new(16, 42)
    $aCodeBox.Controls.Add($aUrlLabel)

    $aStatusLabel = [System.Windows.Forms.Label]::new()
    $aStatusLabel.Text = ""; $aStatusLabel.Font = $FN
    $aStatusLabel.ForeColor = HC '#9CA3AF'; $aStatusLabel.BackColor = [System.Drawing.Color]::Transparent
    $aStatusLabel.Size = [System.Drawing.Size]::new(580, 20); $aStatusLabel.Location = [System.Drawing.Point]::new(20, 226)
    $pAuto.Controls.Add($aStatusLabel)

    $aBtnGetCode = [System.Windows.Forms.Button]::new()
    $aBtnGetCode.Text = "Get Device Code"; $aBtnGetCode.Size = [System.Drawing.Size]::new(160, 34)
    $aBtnGetCode.Location = [System.Drawing.Point]::new(20, 254); $aBtnGetCode.FlatStyle = "Flat"
    $aBtnGetCode.Font = $FB; $aBtnGetCode.BackColor = HC '#7C3AED'; $aBtnGetCode.ForeColor = [System.Drawing.Color]::White
    $aBtnGetCode.Cursor = "Hand"; $aBtnGetCode.FlatAppearance.BorderColor = HC '#4C1D95'; $aBtnGetCode.FlatAppearance.BorderSize = 1
    $pAuto.Controls.Add($aBtnGetCode)

    $aBtnOpenBrowser = [System.Windows.Forms.Button]::new()
    $aBtnOpenBrowser.Text = "Open microsoft.com/devicelogin"; $aBtnOpenBrowser.Size = [System.Drawing.Size]::new(240, 34)
    $aBtnOpenBrowser.Location = [System.Drawing.Point]::new(190, 254); $aBtnOpenBrowser.FlatStyle = "Flat"
    $aBtnOpenBrowser.Font = $FN; $aBtnOpenBrowser.BackColor = HC '#1D4ED8'; $aBtnOpenBrowser.ForeColor = [System.Drawing.Color]::White
    $aBtnOpenBrowser.Cursor = "Hand"; $aBtnOpenBrowser.Enabled = $false
    $aBtnOpenBrowser.FlatAppearance.BorderColor = HC '#1E40AF'; $aBtnOpenBrowser.FlatAppearance.BorderSize = 1
    $pAuto.Controls.Add($aBtnOpenBrowser)

    # ── Manual panel ─────────────────────────────────────────
    $pManual = [System.Windows.Forms.Panel]::new()
    $pManual.SetBounds(0, 104, 620, 316); $pManual.BackColor = HC '#1E1E2E'; $pManual.Visible = $false
    $wiz.Controls.Add($pManual)

    $mInfo = [System.Windows.Forms.Label]::new()
    $mInfo.Text = "Manually enter credentials from an existing Azure AD App Registration. See the documentation for setup instructions."
    $mInfo.Font = $FN; $mInfo.ForeColor = HC '#9CA3AF'; $mInfo.BackColor = [System.Drawing.Color]::Transparent
    $mInfo.Size = [System.Drawing.Size]::new(580, 36); $mInfo.Location = [System.Drawing.Point]::new(20, 14)
    $pManual.Controls.Add($mInfo)

    $mBtnDocs = [System.Windows.Forms.Button]::new()
    $mBtnDocs.Text = "Open Manual Setup Guide"; $mBtnDocs.Size = [System.Drawing.Size]::new(200, 28)
    $mBtnDocs.Location = [System.Drawing.Point]::new(20, 52); $mBtnDocs.FlatStyle = "Flat"
    $mBtnDocs.Font = $FS; $mBtnDocs.BackColor = HC '#374151'; $mBtnDocs.ForeColor = HC '#60A5FA'
    $mBtnDocs.Cursor = "Hand"; $mBtnDocs.FlatAppearance.BorderColor = HC '#4B5563'; $mBtnDocs.FlatAppearance.BorderSize = 1
    $mBtnDocs.add_Click({ Start-Process "https://github.com/jemil/ACLens/blob/main/docs/manual-sp-setup.md" })
    $pManual.Controls.Add($mBtnDocs)

    function New-MField($label, $y, $placeholder) {
        $lbl = [System.Windows.Forms.Label]::new()
        $lbl.Text = $label; $lbl.Font = $FB; $lbl.ForeColor = HC '#E2E8F0'
        $lbl.BackColor = [System.Drawing.Color]::Transparent; $lbl.AutoSize = $true
        $lbl.Location = [System.Drawing.Point]::new(20, $y)
        $pManual.Controls.Add($lbl)
        $tb = [System.Windows.Forms.TextBox]::new()
        $tb.Size = [System.Drawing.Size]::new(578, 26); $tb.Location = [System.Drawing.Point]::new(20, $y + 20)
        $tb.BackColor = HC '#2D2D3F'; $tb.ForeColor = HC '#6B7280'
        $tb.Text = $placeholder; $tb.BorderStyle = "FixedSingle"; $tb.Font = $FN
        $pManual.Controls.Add($tb)
        return $tb
    }

    $mTbTenant = New-MField "Tenant ID"     90  "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
    $mTbClient = New-MField "Client ID"     140 "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
    $mTbSecret = New-MField "Client Secret" 190 "your-client-secret-value"
    $mTbSecret.UseSystemPasswordChar = $false

    $mBtnSave = [System.Windows.Forms.Button]::new()
    $mBtnSave.Text = "Save & Connect"; $mBtnSave.Size = [System.Drawing.Size]::new(160, 34)
    $mBtnSave.Location = [System.Drawing.Point]::new(20, 250); $mBtnSave.FlatStyle = "Flat"
    $mBtnSave.Font = $FB; $mBtnSave.BackColor = HC '#7C3AED'; $mBtnSave.ForeColor = [System.Drawing.Color]::White
    $mBtnSave.Cursor = "Hand"; $mBtnSave.FlatAppearance.BorderColor = HC '#4C1D95'; $mBtnSave.FlatAppearance.BorderSize = 1
    $pManual.Controls.Add($mBtnSave)

    $mBtnSave.add_Click({
        $t = $mTbTenant.Text.Trim(); $c = $mTbClient.Text.Trim(); $sec = $mTbSecret.Text.Trim()
        if (-not $t -or -not $c -or -not $sec -or $t -like "*xxxx*") {
            [System.Windows.Forms.MessageBox]::Show("Please fill in all fields.", "Missing data", "OK", "Warning") | Out-Null
            return
        }
        Save-SPCredentials $t $c $sec
        Update-SPStatus
        [System.Windows.Forms.MessageBox]::Show("Credentials saved successfully.", "ACLens", "OK", "Information") | Out-Null
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
        $aStatusLabel.Text   = "Requesting device code..."
        $aStatusLabel.ForeColor = HC '#9CA3AF'
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
            $aUrlLabel.Text     = "Go to: $($resp.verification_uri)"
            $aStatusLabel.Text  = "Waiting for sign-in... (expires in $([int]($resp.expires_in / 60)) minutes)"
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
                    $aStatusLabel.Text = "Signed in! Creating App Registration..."
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

                    $aStatusLabel.Text = "✔  App Registration created! Client ID: $($app.appId)"
                    $aStatusLabel.ForeColor = HC '#34D399'
                    $aCodeLabel.Text = "Setup complete!"
                    $aCodeLabel.ForeColor = HC '#34D399'

                    Start-Sleep -Milliseconds 1500
                    $wiz.Close()

                } catch {
                    $err = $_.Exception.Message
                    if ($err -notlike "*authorization_pending*" -and $err -notlike "*slow_down*") {
                        if ($err -like "*authorization_declined*" -or $err -like "*expired*") {
                            $script:pollTimer.Stop()
                            $aStatusLabel.Text = "Sign-in cancelled or expired. Try again."
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
    $wSepFtr.SetBounds(0, 420, 620, 1); $wSepFtr.BackColor = HC '#4B5563'
    $wiz.Controls.Add($wSepFtr)

    $wBtnClose = [System.Windows.Forms.Button]::new()
    $wBtnClose.Text = "Close"; $wBtnClose.Size = [System.Drawing.Size]::new(100, 34)
    $wBtnClose.Location = [System.Drawing.Point]::new(500, 434); $wBtnClose.FlatStyle = "Flat"
    $wBtnClose.Font = $FN; $wBtnClose.BackColor = HC '#374151'; $wBtnClose.ForeColor = HC '#E2E8F0'
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

v0.2.1-alpha  (2026-03-27)
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
    $dlg.Size            = [System.Drawing.Size]::new(520, 420)
    $dlg.StartPosition   = "CenterParent"
    $dlg.FormBorderStyle = "FixedDialog"
    $dlg.MaximizeBox     = $false
    $dlg.BackColor       = HC '#1E1E2E'

    $txtLog = [System.Windows.Forms.RichTextBox]::new()
    $txtLog.Dock        = "Fill"
    $txtLog.Text        = $changelogText
    $txtLog.ReadOnly    = $true
    $txtLog.BackColor   = HC '#1A1A2E'
    $txtLog.ForeColor   = HC '#E2E8F0'
    $txtLog.Font        = [System.Drawing.Font]::new("Consolas", 9)
    $txtLog.BorderStyle = "None"
    $txtLog.Padding     = [System.Windows.Forms.Padding]::new(12)
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
    $cTabNTFS.Location = [System.Drawing.Point]::new(16, 62); $cTabNTFS.FlatStyle = "Flat"
    $cTabNTFS.Font = $FB; $cTabNTFS.BackColor = HC '#7C3AED'; $cTabNTFS.ForeColor = [System.Drawing.Color]::White
    $cTabNTFS.Cursor = "Hand"; $cTabNTFS.FlatAppearance.BorderColor = HC '#4C1D95'; $cTabNTFS.FlatAppearance.BorderSize = 1
    $cWin.Controls.Add($cTabNTFS)

    $cTabSP = [System.Windows.Forms.Button]::new()
    $cTabSP.Text = "  SharePoint Compare"; $cTabSP.Size = [System.Drawing.Size]::new(180, 34)
    $cTabSP.Location = [System.Drawing.Point]::new(174, 62); $cTabSP.FlatStyle = "Flat"
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
    $nBtnJson.Location = [System.Drawing.Point]::new(554, 36); $nBtnJson.FlatStyle = "Flat"
    $nBtnJson.BackColor = HC '#374151'; $nBtnJson.ForeColor = HC '#E2E8F0'; $nBtnJson.Cursor = "Hand"
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
    $nBtnPath.Location = [System.Drawing.Point]::new(554, 96); $nBtnPath.FlatStyle = "Flat"
    $nBtnPath.BackColor = HC '#374151'; $nBtnPath.ForeColor = HC '#E2E8F0'; $nBtnPath.Cursor = "Hand"
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
    $sBtnJson.Location = [System.Drawing.Point]::new(554, 36); $sBtnJson.FlatStyle = "Flat"
    $sBtnJson.BackColor = HC '#374151'; $sBtnJson.ForeColor = HC '#E2E8F0'; $sBtnJson.Cursor = "Hand"
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
    $btnRun.Location = [System.Drawing.Point]::new(16, 440); $btnRun.FlatStyle = "Flat"
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


