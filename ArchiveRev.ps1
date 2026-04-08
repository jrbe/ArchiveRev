# ArchiveRev.ps1 v1.2.0
# Archive files or folder contents to an Archive subfolder with a revision suffix.
# Supports open/locked files via 7-Zip shadow copy (-ssw).
# Multi-file: select any number of files, right-click one -> one prompt -> one zip.
# https://github.com/jrbe/ArchiveRev

param(
    [string]$TargetPath,
    [ValidateSet("File","Folder","FolderBackground")]
    [string]$Mode = "File",
    [switch]$Quick
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$logFile = Join-Path $env:LOCALAPPDATA "ArchiveRev\ArchiveRev_error.log"
function Write-Log { param([string]$Msg)
    try { Add-Content -Path $logFile -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $Msg" } catch {}
}

Write-Log "Started. Mode=$Mode Quick=$Quick TargetPath=$TargetPath"

try {

# Hide the PowerShell console window immediately so it doesn't linger on screen.
# WinForms requires the process to have a normal window style at launch (so Windows
# will render the form), but we don't want the console visible to the user.
Add-Type -Name ConsoleHider -Namespace ArchiveRev -MemberDefinition '
    [System.Runtime.InteropServices.DllImport("kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();
    [System.Runtime.InteropServices.DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
' -ErrorAction SilentlyContinue
try {
    $consoleHwnd = [ArchiveRev.ConsoleHider]::GetConsoleWindow()
    [ArchiveRev.ConsoleHider]::ShowWindow($consoleHwnd, 0) | Out-Null
} catch {}
Write-Log "Console hidden."

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ---------------------------------------------------------------------------
# Multi-file queue (File mode only)
# When multiple files are selected, Explorer fires one process per file.
# The first process becomes coordinator and waits for the rest to check in,
# then shows a single dialog for the whole batch.
# ---------------------------------------------------------------------------
$allPaths    = @()
$primaryPath = ""

if ($Mode -eq "File") {

    $queueFile = Join-Path $env:TEMP "ArchiveRev_queue.txt"
    $coordFlag = Join-Path $env:TEMP "ArchiveRev_coord.flag"
    $mutexName = "Local\ArchiveRev_Queue"
    $waitMs    = 650   # ms to wait for sibling processes to check in

    $mutex   = New-Object System.Threading.Mutex($false, $mutexName)
    $isCoord = $false

    try {
        $mutex.WaitOne(5000) | Out-Null
        $isCoord = -not (Test-Path $coordFlag)
        Add-Content -Path $queueFile -Value $TargetPath
        if ($isCoord) { Set-Content -Path $coordFlag -Value $PID }
    } finally {
        try { $mutex.ReleaseMutex() } catch {}
    }

    if (-not $isCoord) {
        Write-Log "Non-coordinator, exiting"
        exit 0
    }

    Write-Log "Coordinator. Waiting ${waitMs}ms for siblings..."
    Start-Sleep -Milliseconds $waitMs

    try {
        $mutex.WaitOne(5000) | Out-Null
        $queued = @(Get-Content $queueFile -ErrorAction SilentlyContinue | Where-Object { $_ })
        Remove-Item $queueFile  -Force -ErrorAction SilentlyContinue
        Remove-Item $coordFlag  -Force -ErrorAction SilentlyContinue
    } finally {
        try { $mutex.ReleaseMutex() } catch {}
    }

    $allPaths    = @($queued)
    $primaryPath = ($allPaths[0]).Trim('"')
    Write-Log "Queue: $($allPaths.Count) files. Primary: $primaryPath"

} else {
    $primaryPath = $TargetPath.TrimEnd('\').Trim('"')
    $allPaths    = @($primaryPath)
    Write-Log "Folder mode. Path: $primaryPath"
}

if (-not $primaryPath -or -not (Test-Path -LiteralPath $primaryPath)) {
    Write-Log "Primary path invalid or not found: $primaryPath"
    [System.Windows.Forms.MessageBox]::Show(
        "Target path not found:`n$primaryPath", "ArchiveRev", "OK", "Error") | Out-Null
    exit 1
}

# ---------------------------------------------------------------------------
# Determine archive dir and default zip base name
# ---------------------------------------------------------------------------
if ($Mode -eq "File") {
    $parentDir   = Split-Path $primaryPath -Parent
    $archiveDir  = Join-Path $parentDir "Archive"
    $defaultName = [System.IO.Path]::GetFileNameWithoutExtension($primaryPath)
} else {
    $archiveDir  = Join-Path $primaryPath "Archive"
    $defaultName = Split-Path $primaryPath -Leaf
}

# ---------------------------------------------------------------------------
# Quick Zip mode - skip dialog, use filename_date.zip directly
# ---------------------------------------------------------------------------
$today = Get-Date -Format "yyyy-MM-dd"

if ($Quick) {
    $script:zipResult = $defaultName + "_" + $today + ".zip"
    Write-Log "Quick mode: $script:zipResult"
} else {

# ---------------------------------------------------------------------------
# WinForms dialog (full mode)
# ---------------------------------------------------------------------------
Write-Log "Loading WinForms assemblies..."
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Write-Log "Assemblies loaded. Building form..."

$form = New-Object System.Windows.Forms.Form
$form.Text          = "ArchiveRev"
$form.ClientSize    = New-Object System.Drawing.Size(420, 340)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox   = $false
$form.MinimizeBox   = $false
$form.BackColor     = [System.Drawing.Color]::White
$form.Font          = New-Object System.Drawing.Font("Segoe UI", 9)
$form.TopMost       = $true
$form.ShowInTaskbar = $true

$pad = 18
$y   = $pad

# --- Header: file count + truncated path ---
$fileCountStr = if ($allPaths.Count -eq 1) { "1 file" } else { "$($allPaths.Count) files selected" }
$lblCount = New-Object System.Windows.Forms.Label
$lblCount.Text     = $fileCountStr
$lblCount.Font     = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$lblCount.Location = New-Object System.Drawing.Point($pad, $y)
$lblCount.Size     = New-Object System.Drawing.Size(384, 18)
$form.Controls.Add($lblCount)
$y += 20

$shortPath = $primaryPath
if ($shortPath.Length -gt 60) { $shortPath = "..." + $shortPath.Substring($shortPath.Length - 57) }
if ($allPaths.Count -gt 1)    { $shortPath += "  (+ $($allPaths.Count - 1) more)" }
$lblPath = New-Object System.Windows.Forms.Label
$lblPath.Text      = $shortPath
$lblPath.ForeColor = [System.Drawing.Color]::Gray
$lblPath.Font      = New-Object System.Drawing.Font("Segoe UI", 8)
$lblPath.Location  = New-Object System.Drawing.Point($pad, $y)
$lblPath.Size      = New-Object System.Drawing.Size(384, 16)
$form.Controls.Add($lblPath)
$y += 24

# Separator
$sep = New-Object System.Windows.Forms.Label
$sep.BorderStyle = "Fixed3D"
$sep.Location    = New-Object System.Drawing.Point($pad, $y)
$sep.Size        = New-Object System.Drawing.Size(384, 2)
$form.Controls.Add($sep)
$y += 12

# --- Zip filename ---
$lblName = New-Object System.Windows.Forms.Label
$lblName.Text     = "Zip filename"
$lblName.Location = New-Object System.Drawing.Point($pad, $y)
$lblName.Size     = New-Object System.Drawing.Size(200, 18)
$form.Controls.Add($lblName)
$y += 20

$tbName = New-Object System.Windows.Forms.TextBox
$tbName.Text     = $defaultName
$tbName.Location = New-Object System.Drawing.Point($pad, $y)
$tbName.Size     = New-Object System.Drawing.Size(384, 24)
$form.Controls.Add($tbName)
$y += 26

$lblNameHint = New-Object System.Windows.Forms.Label
$lblNameHint.Text      = "From right-clicked file. Edit freely."
$lblNameHint.Font      = New-Object System.Drawing.Font("Segoe UI", 8)
$lblNameHint.ForeColor = [System.Drawing.Color]::Gray
$lblNameHint.Location  = New-Object System.Drawing.Point($pad, $y)
$lblNameHint.Size      = New-Object System.Drawing.Size(384, 16)
$form.Controls.Add($lblNameHint)
$y += 22

# --- Revision suffix ---
$lblRev = New-Object System.Windows.Forms.Label
$lblRev.Text     = "Revision suffix"
$lblRev.Location = New-Object System.Drawing.Point($pad, $y)
$lblRev.Size     = New-Object System.Drawing.Size(200, 18)
$form.Controls.Add($lblRev)
$y += 20

$tbRev = New-Object System.Windows.Forms.TextBox
$tbRev.Text     = ""
$tbRev.Location = New-Object System.Drawing.Point($pad, $y)
$tbRev.Size     = New-Object System.Drawing.Size(384, 24)
$form.Controls.Add($tbRev)
$tbRev.Focus() | Out-Null
$y += 30

# --- Date checkbox ---
$cbDate = New-Object System.Windows.Forms.CheckBox
$cbDate.Text     = "Append today's date    $today    (YYYY-MM-DD)"
$cbDate.Checked  = $true
$cbDate.Location = New-Object System.Drawing.Point($pad, $y)
$cbDate.Size     = New-Object System.Drawing.Size(384, 22)
$form.Controls.Add($cbDate)
$y += 30

# --- Live preview panel ---
$previewPanel = New-Object System.Windows.Forms.Panel
$previewPanel.BackColor   = [System.Drawing.Color]::FromArgb(245, 245, 248)
$previewPanel.BorderStyle = "FixedSingle"
$previewPanel.Location    = New-Object System.Drawing.Point($pad, $y)
$previewPanel.Size        = New-Object System.Drawing.Size(384, 40)
$form.Controls.Add($previewPanel)

$lblPreviewCaption = New-Object System.Windows.Forms.Label
$lblPreviewCaption.Text      = "Output  ->  Archive\"
$lblPreviewCaption.Font      = New-Object System.Drawing.Font("Segoe UI", 7.5)
$lblPreviewCaption.ForeColor = [System.Drawing.Color]::Gray
$lblPreviewCaption.Location  = New-Object System.Drawing.Point(8, 4)
$lblPreviewCaption.Size      = New-Object System.Drawing.Size(368, 14)
$previewPanel.Controls.Add($lblPreviewCaption)

$lblPreview = New-Object System.Windows.Forms.Label
$lblPreview.Font     = New-Object System.Drawing.Font("Consolas", 8.5)
$lblPreview.Location = New-Object System.Drawing.Point(8, 20)
$lblPreview.Size     = New-Object System.Drawing.Size(368, 16)
$previewPanel.Controls.Add($lblPreview)

$y += 48

# --- Buttons ---
$btnCreate = New-Object System.Windows.Forms.Button
$btnCreate.Text      = "Create Zip Archive"
$btnCreate.Location  = New-Object System.Drawing.Point($pad, $y)
$btnCreate.Size      = New-Object System.Drawing.Size(280, 30)
$btnCreate.BackColor = [System.Drawing.Color]::FromArgb(28, 28, 28)
$btnCreate.ForeColor = [System.Drawing.Color]::White
$btnCreate.FlatStyle = "Flat"
$btnCreate.FlatAppearance.BorderSize = 0
$form.Controls.Add($btnCreate)
$form.AcceptButton = $btnCreate

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text      = "Cancel"
$btnCancel.Location  = New-Object System.Drawing.Point(310, $y)
$btnCancel.Size      = New-Object System.Drawing.Size(92, 30)
$btnCancel.FlatStyle = "Flat"
$form.Controls.Add($btnCancel)
$form.CancelButton = $btnCancel

# --- Live preview update ---
function Update-Preview {
    $n     = $tbName.Text.Trim()
    $r     = ($tbRev.Text.Trim()) -replace '[\\/:*?"<>|]', '_'
    $parts = @()
    if ($n) { $parts += $n }
    if ($r) { $parts += $r }
    if ($cbDate.Checked) { $parts += $today }
    $lblPreview.Text = if ($parts.Count) { ($parts -join "_") + ".zip" } else { "" }
}

Update-Preview
$tbName.Add_TextChanged({  Update-Preview })
$tbRev.Add_TextChanged({   Update-Preview })
$cbDate.Add_CheckedChanged({ Update-Preview })

# --- Button handlers ---
$script:zipResult = $null

$btnCreate.Add_Click({
    $n = $tbName.Text.Trim()
    $r = ($tbRev.Text.Trim()) -replace '[\\/:*?"<>|]', '_'
    if (-not $n) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please enter a zip filename.", "ArchiveRev", "OK", "Warning") | Out-Null
        return
    }
    $parts = @($n)
    if ($r) { $parts += $r }
    if ($cbDate.Checked) { $parts += $today }
    $script:zipResult = ($parts -join "_") + ".zip"
    $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Close()
})

$btnCancel.Add_Click({ $form.Close() })
Write-Log "Calling ShowDialog..."
$form.Add_Shown({ $form.Activate() })
$form.ShowDialog() | Out-Null
Write-Log "ShowDialog returned. Result: $script:zipResult"

if (-not $script:zipResult) { Write-Log "User cancelled"; exit 0 }

} # end if $Quick / else

Write-Log "Zip name: $script:zipResult"

# ---------------------------------------------------------------------------
# Resolve sources to zip
# ---------------------------------------------------------------------------
if ($Mode -eq "File") {
    $zipSources = @($allPaths | ForEach-Object { $_.Trim('"') })
} else {
    $zipSources = @(Get-ChildItem -LiteralPath $primaryPath |
                    Where-Object { $_.Name -ne "Archive" } |
                    ForEach-Object { $_.FullName })
}

if ($zipSources.Count -eq 0) {
    [System.Windows.Forms.MessageBox]::Show(
        "Nothing to archive (folder is empty or contains only 'Archive').",
        "ArchiveRev", "OK", "Information") | Out-Null
    exit 0
}

# ---------------------------------------------------------------------------
# Create Archive folder and check for overwrite
# ---------------------------------------------------------------------------
if (-not (Test-Path $archiveDir)) {
    New-Item -ItemType Directory -Path $archiveDir | Out-Null
    Write-Log "Created: $archiveDir"
}

$zipPath = Join-Path $archiveDir $script:zipResult

if (Test-Path -LiteralPath $zipPath) {
    $ow = [System.Windows.Forms.MessageBox]::Show(
        "'$script:zipResult' already exists in Archive.`n`nOverwrite?",
        "ArchiveRev", "YesNo", "Question")
    if ($ow -ne "Yes") { Write-Log "Overwrite declined"; exit 0 }
    Remove-Item -LiteralPath $zipPath -Force
}

# ---------------------------------------------------------------------------
# Compress
# ---------------------------------------------------------------------------
$sevenZip = @(
    "${env:ProgramFiles}\7-Zip\7z.exe",
    "${env:ProgramFiles(x86)}\7-Zip\7z.exe"
) | Where-Object { Test-Path $_ } | Select-Object -First 1

if (-not $sevenZip) {
    $found = Get-Command "7z.exe" -ErrorAction SilentlyContinue
    if ($found) { $sevenZip = $found.Source }
}

Write-Log "7-Zip: $sevenZip   Sources: $($zipSources.Count)"

$success = $false
$detail  = ""

if ($sevenZip) {
    $output  = & $sevenZip a -tzip -ssw -mx=5 "$zipPath" @zipSources 2>&1
    $detail  = $output -join "`n"
    $success = ($LASTEXITCODE -le 1)
    Write-Log "7-Zip exit: $LASTEXITCODE"
} else {
    try {
        Compress-Archive -LiteralPath $zipSources -DestinationPath $zipPath -CompressionLevel Optimal
        $success = $true
    } catch {
        $detail = $_.Exception.Message
    }
}

# ---------------------------------------------------------------------------
# Result
# ---------------------------------------------------------------------------
if ($success -and (Test-Path -LiteralPath $zipPath)) {
    $sizeMB = [math]::Round((Get-Item $zipPath).Length / 1MB, 2)
    Write-Log "Success: $zipPath ($sizeMB MB)"
    [System.Windows.Forms.MessageBox]::Show(
        "Archived successfully!`n`n$zipPath`n`n$($zipSources.Count) file(s)   |   ${sizeMB} MB",
        "ArchiveRev", "OK", "Information") | Out-Null
} else {
    $hint = if (-not $sevenZip) {
        "`n`nTip: Install 7-Zip for open-file (shadow copy) support.`nhttps://www.7-zip.org"
    } else { "" }
    Write-Log "FAILED: $detail"
    [System.Windows.Forms.MessageBox]::Show(
        "Archive failed.$hint`n`nDetails:`n$detail`n`nLog: $logFile",
        "ArchiveRev", "OK", "Error") | Out-Null
}

} catch {
    $errMsg  = $_.Exception.Message
    $errLine = $_.InvocationInfo.ScriptLineNumber
    Write-Log "UNHANDLED at line ${errLine}: $errMsg"
    try {
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.MessageBox]::Show(
            "Unexpected error at line ${errLine}:`n$errMsg`n`nLog: $logFile",
            "ArchiveRev Error", "OK", "Error") | Out-Null
    } catch {}
}