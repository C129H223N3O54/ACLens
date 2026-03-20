#Requires -Version 5.1

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

    # Fensterhoehe: Header(56) + Inhalt(140) + Separator(1) + Footer(70) = 267 + Titelleiste(~32)
    $adm = [System.Windows.Forms.Form]::new()
    $adm.Text            = "ACLens"
    $adm.ClientSize      = [System.Drawing.Size]::new(500, 267)
    $adm.StartPosition   = "CenterScreen"
    $adm.FormBorderStyle = "FixedDialog"
    $adm.MaximizeBox     = $false
    $adm.MinimizeBox     = $false
    $adm.BackColor       = HAC '#1E1E2E'
    $adm.Font            = [System.Drawing.Font]::new("Segoe UI", 9)

    # ── Header (0..55) ──────────────────────────────────────
    $aHdr = [System.Windows.Forms.Panel]::new()
    $aHdr.SetBounds(0, 0, 500, 56)
    $aHdr.BackColor = HAC '#1A1A2E'
    $adm.Controls.Add($aHdr)

    $aTitle = [System.Windows.Forms.Label]::new()
    $aTitle.Text      = "ACLens"
    $aTitle.Font      = [System.Drawing.Font]::new("Segoe UI", 15, [System.Drawing.FontStyle]::Bold)
    $aTitle.ForeColor = HAC '#A78BFA'
    $aTitle.BackColor = [System.Drawing.Color]::Transparent
    $aTitle.AutoSize  = $true
    $aTitle.Location  = [System.Drawing.Point]::new(18, 13)
    $aHdr.Controls.Add($aTitle)

    $aSub = [System.Windows.Forms.Label]::new()
    $aSub.Text      = "Administrator Recommended"
    $aSub.Font      = [System.Drawing.Font]::new("Segoe UI", 9)
    $aSub.ForeColor = HAC '#6B7280'
    $aSub.BackColor = [System.Drawing.Color]::Transparent
    $aSub.AutoSize  = $true
    $aSub.Location  = [System.Drawing.Point]::new(130, 21)
    $aHdr.Controls.Add($aSub)

    # Header separator
    $aHdrSep = [System.Windows.Forms.Panel]::new()
    $aHdrSep.SetBounds(0, 56, 500, 1)
    $aHdrSep.BackColor = HAC '#4B5563'
    $adm.Controls.Add($aHdrSep)

    # ── Content (57..196) ───────────────────────────────────
    # Warning icon
    $aIcon = [System.Windows.Forms.Label]::new()
    $aIcon.Text      = [char]0x26A0
    $aIcon.Font      = [System.Drawing.Font]::new("Segoe UI", 28)
    $aIcon.ForeColor = HAC '#FBBF24'
    $aIcon.BackColor = [System.Drawing.Color]::Transparent
    $aIcon.AutoSize  = $true
    $aIcon.Location  = [System.Drawing.Point]::new(18, 66)
    $adm.Controls.Add($aIcon)

    # Main message
    $aMsg1 = [System.Windows.Forms.Label]::new()
    $aMsg1.Text      = "ACLens is not running as Administrator."
    $aMsg1.Font      = [System.Drawing.Font]::new("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $aMsg1.ForeColor = HAC '#E2E8F0'
    $aMsg1.BackColor = [System.Drawing.Color]::Transparent
    $aMsg1.AutoSize  = $true
    $aMsg1.Location  = [System.Drawing.Point]::new(68, 68)
    $adm.Controls.Add($aMsg1)

    $aMsg2 = [System.Windows.Forms.Label]::new()
    $aMsg2.Text      = "Without administrator privileges some folders may not be readable and permission results could be incomplete."
    $aMsg2.Font      = [System.Drawing.Font]::new("Segoe UI", 9)
    $aMsg2.ForeColor = HAC '#9CA3AF'
    $aMsg2.BackColor = [System.Drawing.Color]::Transparent
    $aMsg2.Size      = [System.Drawing.Size]::new(414, 42)
    $aMsg2.Location  = [System.Drawing.Point]::new(68, 96)
    $adm.Controls.Add($aMsg2)

    $aMsg3 = [System.Windows.Forms.Label]::new()
    $aMsg3.Text      = "Restart ACLens as Administrator now?"
    $aMsg3.Font      = [System.Drawing.Font]::new("Segoe UI", 9)
    $aMsg3.ForeColor = HAC '#E2E8F0'
    $aMsg3.BackColor = [System.Drawing.Color]::Transparent
    $aMsg3.AutoSize  = $true
    $aMsg3.Location  = [System.Drawing.Point]::new(68, 148)
    $adm.Controls.Add($aMsg3)

    # ── Footer separator + buttons (197..267) ───────────────
    $aSep = [System.Windows.Forms.Panel]::new()
    $aSep.SetBounds(0, 196, 500, 1)
    $aSep.BackColor = HAC '#4B5563'
    $adm.Controls.Add($aSep)

    $aYes = [System.Windows.Forms.Button]::new()
    $aYes.Text      = "Yes, restart as Admin"
    $aYes.Size      = [System.Drawing.Size]::new(180, 38)
    $aYes.Location  = [System.Drawing.Point]::new(18, 214)
    $aYes.FlatStyle = "Flat"
    $aYes.Font      = [System.Drawing.Font]::new("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $aYes.BackColor = HAC '#7C3AED'
    $aYes.ForeColor = [System.Drawing.Color]::White
    $aYes.Cursor    = "Hand"
    $aYes.FlatAppearance.BorderColor = HAC '#4C1D95'
    $aYes.FlatAppearance.BorderSize  = 1
    $aYes.DialogResult = [System.Windows.Forms.DialogResult]::Yes
    $adm.Controls.Add($aYes)

    $aNo = [System.Windows.Forms.Button]::new()
    $aNo.Text      = "Continue without Admin"
    $aNo.Size      = [System.Drawing.Size]::new(180, 38)
    $aNo.Location  = [System.Drawing.Point]::new(208, 214)
    $aNo.FlatStyle = "Flat"
    $aNo.Font      = [System.Drawing.Font]::new("Segoe UI", 9)
    $aNo.BackColor = HAC '#374151'
    $aNo.ForeColor = HAC '#E2E8F0'
    $aNo.Cursor    = "Hand"
    $aNo.FlatAppearance.BorderColor = HAC '#4B5563'
    $aNo.FlatAppearance.BorderSize  = 1
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

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Web

# ============================================================
# HILFSFUNKTIONEN (NTFS-Logik)
# ============================================================

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

# ============================================================
# GUI
# ============================================================

[System.Windows.Forms.Application]::EnableVisualStyles()

function HC([string]$h) {
    $h = $h.TrimStart('#')
    [System.Drawing.Color]::FromArgb(
        [Convert]::ToInt32($h.Substring(0,2),16),
        [Convert]::ToInt32($h.Substring(2,2),16),
        [Convert]::ToInt32($h.Substring(4,2),16))
}

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
$form.MinimumSize     = [System.Drawing.Size]::new(600, 420)
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
$lblTitle.Location  = [System.Drawing.Point]::new(16, 12)
$pnlHdr.Controls.Add($lblTitle)

$lblSub = [System.Windows.Forms.Label]::new()
$lblSub.Text      = "NTFS Permission Analyzer & Reporter"
$lblSub.Font      = $FS
$lblSub.ForeColor = $C_LOW
$lblSub.BackColor = [System.Drawing.Color]::Transparent
$lblSub.AutoSize  = $true
$lblSub.Location  = [System.Drawing.Point]::new(118, 21)
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
$form.Controls.Add($lblPath)
$form.Controls.Add($tbPath)
$form.Controls.Add($btnBrowse)

# Zeile 2: Output File
$lblOut  = New-Label "Output File (HTML + JSON):" $true
$tbOut   = New-TB "(automatic — saved to script folder)"
$btnSave = New-Btn "Save as..."
$form.Controls.Add($lblOut)
$form.Controls.Add($tbOut)
$form.Controls.Add($btnSave)

# Zeile 3: Depth
$lblDepth = New-Label "Maximum Depth:" $true
$form.Controls.Add($lblDepth)

$numDepth = [System.Windows.Forms.NumericUpDown]::new()
$numDepth.Size      = [System.Drawing.Size]::new(68, $BTN_H)
$numDepth.Minimum   = 0
$numDepth.Maximum   = 100
$numDepth.Value     = 0
$numDepth.BackColor = $C_INPUT
$numDepth.ForeColor = $C_TXT
$numDepth.Font      = $FN
$form.Controls.Add($numDepth)

$lblUnlim = New-Label "  0 = unlimited" $false
$lblUnlim.ForeColor = $C_LOW
$form.Controls.Add($lblUnlim)

# Checkbox
$chk = [System.Windows.Forms.CheckBox]::new()
$chk.Text      = "Open report in browser after creation"
$chk.Font      = $FN
$chk.ForeColor = $C_MID
$chk.BackColor = [System.Drawing.Color]::Transparent
$chk.Checked   = $true
$chk.AutoSize  = $true
$form.Controls.Add($chk)

# Separator (zwischen Feldern und Progress)
$sepMid = [System.Windows.Forms.Panel]::new()
$sepMid.BackColor = $C_BORDER
$form.Controls.Add($sepMid)

# ── Progress ─────────────────────────────────────────────────
$lblStatus = [System.Windows.Forms.Label]::new()
$lblStatus.Text      = "Ready."
$lblStatus.Font      = $FN
$lblStatus.ForeColor = $C_MID
$lblStatus.BackColor = [System.Drawing.Color]::Transparent
$form.Controls.Add($lblStatus)

$progBar = [System.Windows.Forms.ProgressBar]::new()
$progBar.Style   = "Continuous"
$progBar.Minimum = 0
$progBar.Maximum = 100
$form.Controls.Add($progBar)

$lblStats = [System.Windows.Forms.Label]::new()
$lblStats.Text      = ""
$lblStats.Font      = $FS
$lblStats.ForeColor = $C_MID
$lblStats.BackColor = [System.Drawing.Color]::Transparent
$form.Controls.Add($lblStats)

# ── Layout-Funktion (absolut, sauber) ────────────────────────
function Do-Layout {
    $cw = $form.ClientSize.Width
    $ch = $form.ClientSize.Height

    $x      = $PAD
    $tbW    = $cw - $PAD * 2 - $BTN_W - 8   # Textbox-Breite
    $fullW  = $cw - $PAD * 2                 # Volle Inhaltsbreite

    # Header & Trennlinie
    $pnlHdr.SetBounds(0, 0, $cw, $HDR_H)
    $sepHdr.SetBounds(0, $HDR_H, $cw, 1)

    # Footer & Trennlinie (von unten)
    $sepFtr.SetBounds(0, $ch - $FTR_H - 1, $cw, 1)
    $pnlFtr.SetBounds(0, $ch - $FTR_H, $cw, $FTR_H)


    # Inhaltsbereich Y-Start
    $contentTop = $HDR_H + 1

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

$form.add_Load({   Do-Layout })
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

$btnOpenLast.add_Click({
    if ($script:lastReport -and (Test-Path $script:lastReport)) {
        Start-Process $script:lastReport
    }
})

$changelogText = @"
ACLens Changelog
================

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

# ── Compare Window ────────────────────────────────────────────
$btnCompare.add_Click({
    $cWin = [System.Windows.Forms.Form]::new()
    $cWin.Text            = "ACLens - Compare Scans"
    $cWin.Size            = [System.Drawing.Size]::new(660, 380)
    $cWin.StartPosition   = "CenterParent"
    $cWin.FormBorderStyle = "FixedDialog"
    $cWin.MaximizeBox     = $false
    $cWin.BackColor       = HC '#1E1E2E'
    $cWin.Font            = $FN

    # Header
    $cHdr = [System.Windows.Forms.Panel]::new()
    $cHdr.Size      = [System.Drawing.Size]::new(660, 52)
    $cHdr.Location  = [System.Drawing.Point]::new(0, 0)
    $cHdr.BackColor = HC '#1A1A2E'
    $cWin.Controls.Add($cHdr)

    $cTitle = [System.Windows.Forms.Label]::new()
    $cTitle.Text      = "Compare Scans"
    $cTitle.Font      = [System.Drawing.Font]::new("Segoe UI", 13, [System.Drawing.FontStyle]::Bold)
    $cTitle.ForeColor = HC '#A78BFA'
    $cTitle.BackColor = [System.Drawing.Color]::Transparent
    $cTitle.AutoSize  = $true
    $cTitle.Location  = [System.Drawing.Point]::new(16, 12)
    $cHdr.Controls.Add($cTitle)

    $cSub = [System.Windows.Forms.Label]::new()
    $cSub.Text      = "Select a baseline JSON and run a new live scan to compare"
    $cSub.Font      = $FS
    $cSub.ForeColor = HC '#6B7280'
    $cSub.BackColor = [System.Drawing.Color]::Transparent
    $cSub.AutoSize  = $true
    $cSub.Location  = [System.Drawing.Point]::new(172, 20)
    $cHdr.Controls.Add($cSub)

    # Baseline JSON
    $lbl1 = [System.Windows.Forms.Label]::new()
    $lbl1.Text = "Baseline (JSON):"; $lbl1.Font = $FB; $lbl1.ForeColor = HC '#E2E8F0'
    $lbl1.BackColor = [System.Drawing.Color]::Transparent; $lbl1.AutoSize = $true
    $lbl1.Location = [System.Drawing.Point]::new(20, 72)
    $cWin.Controls.Add($lbl1)

    $tbJson = [System.Windows.Forms.TextBox]::new()
    $tbJson.Size = [System.Drawing.Size]::new(490, 26); $tbJson.Location = [System.Drawing.Point]::new(20, 94)
    $tbJson.BackColor = HC '#2D2D3F'; $tbJson.ForeColor = HC '#E2E8F0'; $tbJson.BorderStyle = "FixedSingle"
    if ($script:lastJson -and (Test-Path $script:lastJson)) { $tbJson.Text = $script:lastJson }
    $cWin.Controls.Add($tbJson)

    $btnPickJson = [System.Windows.Forms.Button]::new()
    $btnPickJson.Text = "Browse..."; $btnPickJson.Size = [System.Drawing.Size]::new(110, 26)
    $btnPickJson.Location = [System.Drawing.Point]::new(516, 94); $btnPickJson.FlatStyle = "Flat"
    $btnPickJson.BackColor = HC '#374151'; $btnPickJson.ForeColor = HC '#E2E8F0'; $btnPickJson.Cursor = "Hand"
    $btnPickJson.FlatAppearance.BorderColor = HC '#4B5563'; $btnPickJson.FlatAppearance.BorderSize = 1
    $btnPickJson.add_Click({
        $ofd = [System.Windows.Forms.OpenFileDialog]::new()
        $ofd.Filter = "ACLens JSON (*.json)|*.json|All files (*.*)|*.*"
        $ofd.Title  = "Select baseline JSON snapshot"
        if ($ofd.ShowDialog() -eq "OK") { $tbJson.Text = $ofd.FileName }
    })
    $cWin.Controls.Add($btnPickJson)

    # Live Scan Path
    $lbl2 = [System.Windows.Forms.Label]::new()
    $lbl2.Text = "Current Scan Path:"; $lbl2.Font = $FB; $lbl2.ForeColor = HC '#E2E8F0'
    $lbl2.BackColor = [System.Drawing.Color]::Transparent; $lbl2.AutoSize = $true
    $lbl2.Location = [System.Drawing.Point]::new(20, 136)
    $cWin.Controls.Add($lbl2)

    $tbScanPath = [System.Windows.Forms.TextBox]::new()
    $tbScanPath.Size = [System.Drawing.Size]::new(490, 26); $tbScanPath.Location = [System.Drawing.Point]::new(20, 158)
    $tbScanPath.BackColor = HC '#2D2D3F'; $tbScanPath.ForeColor = HC '#E2E8F0'; $tbScanPath.BorderStyle = "FixedSingle"
    if ($tbPath.Text) { $tbScanPath.Text = $tbPath.Text }
    $cWin.Controls.Add($tbScanPath)

    $btnPickPath = [System.Windows.Forms.Button]::new()
    $btnPickPath.Text = "Browse..."; $btnPickPath.Size = [System.Drawing.Size]::new(110, 26)
    $btnPickPath.Location = [System.Drawing.Point]::new(516, 158); $btnPickPath.FlatStyle = "Flat"
    $btnPickPath.BackColor = HC '#374151'; $btnPickPath.ForeColor = HC '#E2E8F0'; $btnPickPath.Cursor = "Hand"
    $btnPickPath.FlatAppearance.BorderColor = HC '#4B5563'; $btnPickPath.FlatAppearance.BorderSize = 1
    $btnPickPath.add_Click({
        $fbd = [System.Windows.Forms.FolderBrowserDialog]::new()
        $fbd.Description = "Select folder to scan"
        if ($tbScanPath.Text -and (Test-Path $tbScanPath.Text)) { $fbd.SelectedPath = $tbScanPath.Text }
        if ($fbd.ShowDialog() -eq "OK") { $tbScanPath.Text = $fbd.SelectedPath }
    })
    $cWin.Controls.Add($btnPickPath)

    # Status
    $cStatus = [System.Windows.Forms.Label]::new()
    $cStatus.Text = "Ready."; $cStatus.Font = $FN; $cStatus.ForeColor = HC '#6B7280'
    $cStatus.BackColor = [System.Drawing.Color]::Transparent
    $cStatus.Size = [System.Drawing.Size]::new(620, 20); $cStatus.Location = [System.Drawing.Point]::new(20, 200)
    $cWin.Controls.Add($cStatus)

    $cBar = [System.Windows.Forms.ProgressBar]::new()
    $cBar.Size = [System.Drawing.Size]::new(620, 12); $cBar.Location = [System.Drawing.Point]::new(20, 226)
    $cBar.Style = "Continuous"; $cBar.Minimum = 0; $cBar.Maximum = 100
    $cWin.Controls.Add($cBar)

    # Separator
    $cSep = [System.Windows.Forms.Panel]::new()
    $cSep.Size = [System.Drawing.Size]::new(660, 1); $cSep.Location = [System.Drawing.Point]::new(0, 258)
    $cSep.BackColor = HC '#4B5563'
    $cWin.Controls.Add($cSep)

    # Run button
    $btnRun = [System.Windows.Forms.Button]::new()
    $btnRun.Text = "Run Comparison"; $btnRun.Size = [System.Drawing.Size]::new(160, 36)
    $btnRun.Location = [System.Drawing.Point]::new(20, 272); $btnRun.FlatStyle = "Flat"
    $btnRun.Font = $FB; $btnRun.BackColor = HC '#7C3AED'; $btnRun.ForeColor = [System.Drawing.Color]::White
    $btnRun.Cursor = "Hand"; $btnRun.FlatAppearance.BorderColor = HC '#4C1D95'; $btnRun.FlatAppearance.BorderSize = 1
    $cWin.Controls.Add($btnRun)

    $btnRun.add_Click({
        $jsonPath = $tbJson.Text.Trim()
        $scanPath = $tbScanPath.Text.Trim()

        if (-not $jsonPath -or -not (Test-Path $jsonPath)) {
            [System.Windows.Forms.MessageBox]::Show("Please select a valid baseline JSON file.", "Missing input", "OK", "Warning") | Out-Null
            return
        }
        if (-not $scanPath -or -not (Test-Path $scanPath -PathType Container)) {
            [System.Windows.Forms.MessageBox]::Show("Please enter a valid scan path.", "Missing input", "OK", "Warning") | Out-Null
            return
        }

        $btnRun.Enabled = $false
        $cBar.Value     = 0
        $cStatus.Text   = "Loading baseline..."
        $cStatus.ForeColor = HC '#9CA3AF'
        $cWin.Refresh()

        try {
            $oldData  = Load-JsonSnapshot -Path $jsonPath
            $oldLabel = $jsonPath

            $cStatus.Text = "Step 1/2: Scanning folders..."
            $cWin.Refresh()
            $allFolders = Get-AllFolders -RootPath $scanPath -MaxDepth 0
            $total = $allFolders.Count

            $newData       = @{}
            $rootPathClean = $scanPath.TrimEnd('\')
            $counter       = 0

            foreach ($fp in $allFolders) {
                $counter++
                $pct = [int](($counter / $total) * 100)
                if ($cBar.Value -ne $pct) { $cBar.Value = $pct }
                if ($counter % 5 -eq 0 -or $counter -eq $total) {
                    $short = if ($fp.Length -gt 80) { "..." + $fp.Substring($fp.Length-77) } else { $fp }
                    $cStatus.Text = "Step 2/2: $counter / $total  -  $short"
                    $cWin.Refresh()
                }
                $aclData = Get-FolderACL -FolderPath $fp
                $newData[$fp] = [PSCustomObject]@{
                    Path                    = $fp
                    Owner                   = $aclData.Owner
                    InheritanceEnabled      = $aclData.InheritanceEnabled
                    AreAccessRulesProtected = $aclData.AreAccessRulesProtected
                    Rules                   = $aclData.Rules
                }
            }

            $cBar.Value   = 99
            $cStatus.Text = "Generating diff report..."
            $cWin.Refresh()

            $ts        = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
            $outPath   = Join-Path $PSScriptRoot "ACLens_Compare_$ts.html"
            $result    = New-CompareHTMLReport -OldData $oldData -NewData $newData -OldLabel $oldLabel -NewLabel $scanPath -OutputPath $outPath

            $cBar.Value        = 100
            $cStatus.Text      = "Done! Changed: $($result.Changed)  Added: $($result.Added)  Removed: $($result.Removed)  - $outPath"
            $cStatus.ForeColor = HC '#34D399'

            Start-Process $outPath
        }
        catch {
            $cStatus.Text      = "Error: $($_.Exception.Message)"
            $cStatus.ForeColor = HC '#F87171'
        }
        finally {
            $btnRun.Enabled = $true
        }
    })

    $cWin.ShowDialog($form) | Out-Null
})

[System.Windows.Forms.Application]::Run($form)
