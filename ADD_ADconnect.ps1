<#
    Azure AD Guest User Inviter – PowerShell 7 edition
    Copyright (c) 2025 Saverio Sacchetti (https://github.com/savershell/)
    Licensed under the MIT License.

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software to use, copy, modify, merge, publish, distribute, sublicense,
    and/or sell copies, subject to the conditions in the LICENSE file.
#>

# ─── 1) Ensure Graph SDK is present ─────────────────────────────────
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Identity.DirectoryManagement)) {
    try {
        Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
    } catch {
        Write-Warning "Could not install Microsoft.Graph: $($_.Exception.Message)"
        return
    }
}
Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop

# ─── 2) WinForms GUI setup ──────────────────────────────────────────
Add-Type -AssemblyName System.Windows.Forms

$form = [Windows.Forms.Form]@{Text='Azure AD Guest User Inviter';Size=[Drawing.Size]::new(600,180);StartPosition='CenterScreen'}
$lblCsv    = [Windows.Forms.Label]@{Text='CSV File:';Location=[Drawing.Point]::new(20,30);AutoSize=$true}
$txtCsv    = [Windows.Forms.TextBox]@{Size=[Drawing.Size]::new(400,20);Location=[Drawing.Point]::new(90,28)}
$btnBrowse = [Windows.Forms.Button]@{Text='Browse';Size=[Drawing.Size]::new(70,23);Location=[Drawing.Point]::new(500,26)}
$btnImport = [Windows.Forms.Button]@{Text='Connect and Add Users';Size=[Drawing.Size]::new(180,30);Location=[Drawing.Point]::new(210,80)}
$btnCancel = [Windows.Forms.Button]@{Text='Cancel';Size=[Drawing.Size]::new(80,30);Location=[Drawing.Point]::new(410,80)}
$form.Controls.AddRange(@($lblCsv,$txtCsv,$btnBrowse,$btnImport,$btnCancel))
$form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon("C:\scripts\eu.ico")

# ─── helper to compress Graph messages ─────────────────────────────
function Clean‑GraphMessage {
    param([string]$Text)
    if (-not $Text) { return $Text }
    $Text = $Text -replace '\[.*?\]\s*:', ''
    $Text = $Text -replace ' as object ID:.*', ''
    ($Text -split '\.')[0].Trim()
}

# ─── Browse button ─────────────────────────────────────────────────
$btnBrowse.Add_Click({
    $dlg = [Windows.Forms.OpenFileDialog]@{Filter='CSV files (*.csv)|*.csv|All files (*.*)|*.*'}
    if ($dlg.ShowDialog() -eq 'OK') { $txtCsv.Text = $dlg.FileName }
})

# ─── Main button: connect and invite guests ────────────────────────
$btnImport.Add_Click({

    # 1) CSV validation
    $csvPath = $txtCsv.Text
    if (-not (Test-Path $csvPath)) { [Windows.Forms.MessageBox]::Show('CSV file not found.'); return }

    try   { $users = Import-Csv -Path $csvPath }
    catch { [Windows.Forms.MessageBox]::Show("Unable to read CSV:`n$_"); return }

    if (-not $users) { [Windows.Forms.MessageBox]::Show('CSV is empty.'); return }
    if (-not ($users[0].PSObject.Properties.Name -contains 'Email')) {
        [Windows.Forms.MessageBox]::Show("CSV must contain an 'Email' column."); return
    }

    if ([Windows.Forms.MessageBox]::Show("Invite $($users.Count) user(s)?",'Confirm','YesNo') -ne 'Yes') { return }

    # 2) Connect to Graph
    $form.Enabled = $false
    try { Connect-MgGraph -Scopes 'User.Invite.All' }
    catch { try { Connect-MgGraph -Scopes 'User.Invite.All' -UseDeviceCode } catch { [Windows.Forms.MessageBox]::Show("Graph sign‑in failed:`n$_"); $form.Enabled=$true; return } }
    $form.Enabled = $true

    # 3) Invitation loop
    $duplicates = @()
    $successes  = @()
    $errors     = @()

    foreach ($u in $users) {

        # Pre‑check: Mail OR UPN
        try {
            $filter = "Mail eq '$($u.Email)' or UserPrincipalName eq '$($u.Email)'"
            $exists = Get-MgUser -Filter $filter -ConsistencyLevel eventual -ErrorAction Stop
        } catch {
            $errors += "$($u.Email)  ➜  lookup failed  ➜  $(Clean‑GraphMessage $_.Exception.Message)"
            continue
        }

        if ($exists) {
            $duplicates += "$($u.Email) (already in tenant)"
            continue
        }

        # Invite
        try {
            $display = $u.DisplayName ? $u.DisplayName : $u.Email
            New-MgInvitation -InvitedUserEmailAddress $u.Email `
                             -InvitedUserDisplayName  $display `
                             -InviteRedirectUrl       'https://myapps.microsoft.com' `
                             -SendInvitationMessage `
                             -ErrorAction Stop
            # success
            $successes += "$($u.Email)"
        } catch {
            $msg = $_.Exception.Message
            if (-not $msg) { $msg = ($_ | Out-String).Split("`n")[0] }

            if ($msg -match 'already exists in the directory') {
                $duplicates += "$($u.Email) (already in tenant)"
            } else {
                $errors += "$($u.Email)  ➜  $(Clean‑GraphMessage $msg)"
            }
        }
    }

    # 4) Summary popup
    [string]$summary = ''
    if ($duplicates) { $summary += "🛈  Duplicates skipped:`n" + ($duplicates -join "`n") + "`n`n" }
    if ($successes)  { $summary += "✓  Invitations sent:`n"    + ($successes  -join "`n") + "`n`n" }
    if ($errors)     { $summary += "⚠️  Problems:`n"          + ($errors     -join "`n") }
    if (-not $summary) { $summary = "✓  No changes needed – every user already exists." }

    [Windows.Forms.MessageBox]::Show($summary.Trim(),'Invitation Summary')
})

$btnCancel.Add_Click({ $form.Close() })
$form.Add_Shown({ $form.Activate() })
[void]$form.ShowDialog()
