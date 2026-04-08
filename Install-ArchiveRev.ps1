# Install-ArchiveRev.ps1 v1.2.0
# Installs ArchiveRev context menu entries for the current user.
# No admin rights required (writes to HKCU only).
# Uses a VBScript launcher so no console window appears when invoked from Explorer.

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName Microsoft.VisualBasic

function Show-Info  { param([string]$Msg) [Microsoft.VisualBasic.Interaction]::MsgBox($Msg, 64, "ArchiveRev Installer") | Out-Null }
function Show-Error { param([string]$Msg) [Microsoft.VisualBasic.Interaction]::MsgBox($Msg, 16, "ArchiveRev Installer - Error") | Out-Null }
function Ask-YesNo  {
    param([string]$Msg, [string]$Title = "ArchiveRev Installer")
    return ([Microsoft.VisualBasic.Interaction]::MsgBox($Msg, 4 + 32, $Title) -eq 6)
}

# ---------------------------------------------------------------------------
# Registry backup
# ---------------------------------------------------------------------------
$backupDir  = Join-Path $env:LOCALAPPDATA "ArchiveRev\Backups"
$backupFile = Join-Path $backupDir ("registry_backup_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".reg")
New-Item -ItemType Directory -Force -Path $backupDir | Out-Null

Write-Host "Backing up registry to:`n  $backupFile"

$keysToBackup = @(
    "HKCU\Software\Classes\*\shell\ArchiveRev",
    "HKCU\Software\Classes\Directory\shell\ArchiveRev",
    "HKCU\Software\Classes\Directory\Background\shell\ArchiveRev",
    "HKCU\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}"
)

"Windows Registry Editor Version 5.00`r`n" | Set-Content -Path $backupFile -Encoding Unicode
foreach ($key in $keysToBackup) {
    try { $null = & reg.exe export $key "$backupFile.tmp" /y 2>&1 } catch {}
    if (Test-Path "$backupFile.tmp") {
        $content = Get-Content "$backupFile.tmp" -Raw
        $content = $content -replace "^Windows Registry Editor Version 5\.00\r?\n", ""
        Add-Content -Path $backupFile -Value $content -Encoding Unicode
        Remove-Item "$backupFile.tmp" -Force
    }
}
Write-Host "  Backup saved."

# ---------------------------------------------------------------------------
# Copy scripts to install location
# ---------------------------------------------------------------------------
$installDir  = Join-Path $env:LOCALAPPDATA "ArchiveRev"
$scriptDest  = Join-Path $installDir "ArchiveRev.ps1"
$launcherDest = Join-Path $installDir "ArchiveRev_Launch.vbs"

$filesToCopy = @("ArchiveRev.ps1", "ArchiveRev_Launch.vbs", "Toggle-ClassicMenu.ps1", "Toggle-ScriptExecution.ps1")

foreach ($f in $filesToCopy) {
    $src = Join-Path $PSScriptRoot $f
    if (-not (Test-Path $src)) {
        Show-Error "$f not found next to this installer.`nMake sure all files are in the same folder."
        exit 1
    }
}

Write-Host "Installing to: $installDir"
New-Item -ItemType Directory -Force -Path $installDir | Out-Null
foreach ($f in $filesToCopy) {
    Copy-Item -Path (Join-Path $PSScriptRoot $f) -Destination (Join-Path $installDir $f) -Force
    Write-Host "  Copied: $f"
}

# ---------------------------------------------------------------------------
# Build launcher command strings
# wscript.exe runs the VBScript invisibly (window style 0).
# ---------------------------------------------------------------------------
function Make-Command {
    param([string]$Mode, [string]$PathMacro, [switch]$Quick)
    $quickArg = if ($Quick) { " Quick" } else { "" }
    return "wscript.exe `"$launcherDest`" `"$scriptDest`" `"$PathMacro`" $Mode$quickArg"
}

$parentLabel = "Archive Rev"
$menuIcon    = "shell32.dll,44"

# ---------------------------------------------------------------------------
# Register cascade menu using .NET Registry API directly.
# IMPORTANT: Both PowerShell (New-Item) and reg.exe expand * as a wildcard,
# hanging for minutes on HKCU\Software\Classes\*. The .NET Registry class
# calls Win32 RegCreateKeyEx directly, treating * as a literal key name.
# ---------------------------------------------------------------------------
$hkcu = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
    [Microsoft.Win32.RegistryHive]::CurrentUser,
    [Microsoft.Win32.RegistryView]::Default
)

function Set-RegKey {
    param([string]$SubKey, [hashtable]$Values)
    $key = $hkcu.CreateSubKey($SubKey)
    foreach ($kv in $Values.GetEnumerator()) {
        $name = if ($kv.Key -eq "(default)") { "" } else { $kv.Key }
        $key.SetValue($name, $kv.Value)
    }
    $key.Close()
}

function Register-CascadeMenu {
    param([string]$ClassPath)

    $base      = "Software\Classes\$ClassPath\ArchiveRev"
    $pathMacro = if ($ClassPath -eq "Directory\Background\shell") { "%V" } else { "%1" }
    $mode      = if     ($ClassPath -match "Background") { "FolderBackground" }
                 elseif ($ClassPath -match "Directory")  { "Folder" }
                 else                                    { "File" }

    # Note: do NOT set (default) on cascade parent - it conflicts with MUIVerb
    # and causes Windows to treat the entry as a direct command instead of cascade
    Set-RegKey $base @{
        "MUIVerb"     = $parentLabel
        "Icon"        = $menuIcon
        "SubCommands" = ""
    }
    $sk = $hkcu.CreateSubKey("$base\shell"); $sk.Close()

    Set-RegKey "$base\shell\01_Full"         @{ "(default)" = "Archive Rev..."; "MUIVerb" = "Archive Rev..." }
    Set-RegKey "$base\shell\01_Full\command" @{ "(default)" = (Make-Command $mode $pathMacro) }

    Set-RegKey "$base\shell\02_Quick"         @{ "(default)" = "Quick Zip"; "MUIVerb" = "Quick Zip" }
    Set-RegKey "$base\shell\02_Quick\command" @{ "(default)" = (Make-Command $mode $pathMacro -Quick) }

    Write-Host "  Registered: $ClassPath  [$mode]"
}

Write-Host "`nRegistering context menu entries..."
Register-CascadeMenu "*\shell"
Register-CascadeMenu "Directory\shell"
Register-CascadeMenu "Directory\Background\shell"
$hkcu.Close()
# ---------------------------------------------------------------------------
# Offer to enable classic right-click menu on Windows 11
# ---------------------------------------------------------------------------
$regClassic       = "HKCU:\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32"
$classicAlreadyOn = Test-Path $regClassic
$enabledClassic   = $false

if (-not $classicAlreadyOn) {
    $enabledClassic = Ask-YesNo (
        "Windows 11 hides extra context menu items behind 'Show more options'.`n`n" +
        "Would you like to enable the classic right-click menu so " +
        "'Archive Rev' appears immediately - no extra click required?`n`n" +
        "Per-user only, reversible. Run Toggle-ClassicMenu.ps1 to switch back."
    ) "Enable Classic Right-Click Menu?"

    if ($enabledClassic) {
        New-Item -Path $regClassic -Force | Out-Null
        Set-ItemProperty -Path $regClassic -Name "(default)" -Value "" -Force
        Write-Host "  Classic menu enabled."
        Write-Host "  Restarting Explorer..."
        Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
        Start-Sleep -Milliseconds 800
        Start-Process explorer
    }
} else {
    Write-Host "  Classic menu already active."
}

# ---------------------------------------------------------------------------
# Done
# ---------------------------------------------------------------------------
$classicNote = if ($enabledClassic -or $classicAlreadyOn) {
    "Right-click any file or folder -> Archive Rev -> Archive Rev... or Quick Zip"
} else {
    "On Windows 11, right-click -> Show more options -> Archive Rev`nRun Toggle-ClassicMenu.ps1 to show it directly."
}

Show-Info (
    "ArchiveRev v1.2.0 installed!`n`n" +
    $classicNote + "`n`n" +
    "Registry backup:`n$backupFile`n`n" +
    "Tip: Install 7-Zip for open-file (shadow copy) support.`nhttps://www.7-zip.org"
)

Write-Host "`nInstall complete."