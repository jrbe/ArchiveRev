# Uninstall-ArchiveRev.ps1 v1.2.0
# Removes ArchiveRev context menu entries and installed files.
# Optionally restores the Windows 11 modern right-click menu if classic was enabled.

Set-StrictMode -Version Latest
$ErrorActionPreference = "SilentlyContinue"

Add-Type -AssemblyName Microsoft.VisualBasic

function Ask-YesNo {
    param([string]$Msg, [string]$Title = "ArchiveRev Uninstaller")
    $r = [Microsoft.VisualBasic.Interaction]::MsgBox($Msg, 4 + 32, $Title)
    return ($r -eq 6)
}

# ---------------------------------------------------------------------------
# Back up before removing anything
# ---------------------------------------------------------------------------
$backupDir  = Join-Path $env:LOCALAPPDATA "ArchiveRev\Backups"
$backupFile = Join-Path $backupDir ("registry_backup_uninstall_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".reg")
New-Item -ItemType Directory -Force -Path $backupDir | Out-Null

Write-Host "Backing up registry before uninstall to:`n  $backupFile"

$keysToBackup = @(
    "HKCU\Software\Classes\*\shell\ArchiveRev",
    "HKCU\Software\Classes\Directory\shell\ArchiveRev",
    "HKCU\Software\Classes\Directory\Background\shell\ArchiveRev",
    "HKCU\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}"
)

"Windows Registry Editor Version 5.00`r`n" | Set-Content -Path $backupFile -Encoding Unicode
foreach ($key in $keysToBackup) {
    try {
        $null = & reg.exe export $key "$backupFile.tmp" /y 2>&1
    } catch { }
    if (Test-Path "$backupFile.tmp") {
        $content = Get-Content "$backupFile.tmp" -Raw
        $content = $content -replace "^Windows Registry Editor Version 5\.00\r?\n", ""
        Add-Content -Path $backupFile -Value $content -Encoding Unicode
        Remove-Item "$backupFile.tmp" -Force
    }
}

# ---------------------------------------------------------------------------
# Remove context menu entries
# ---------------------------------------------------------------------------
Write-Host "`nRemoving context menu entries..."

# Use reg.exe to avoid PowerShell wildcard expansion on the * path
$verbKeys = @(
    "HKCU\Software\Classes\*\shell\ArchiveRev",
    "HKCU\Software\Classes\Directory\shell\ArchiveRev",
    "HKCU\Software\Classes\Directory\Background\shell\ArchiveRev"
)

foreach ($key in $verbKeys) {
    $null = & reg.exe delete $key /f 2>&1
    Write-Host "  Removed: $key"
}

# ---------------------------------------------------------------------------
# Offer to revert classic right-click menu (if it was enabled)
# ---------------------------------------------------------------------------
$regClassic = "HKCU:\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32"

if (Test-Path $regClassic) {
    $revert = Ask-YesNo (
        "The classic right-click menu (Windows 10 style) is currently enabled.`n`n" +
        "Would you like to revert to the Windows 11 default " +
        "('Show more options') menu as well?"
    ) "Revert Right-Click Menu?"

    if ($revert) {
        Remove-Item -Path "HKCU:\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}" `
                    -Recurse -Force
        Write-Host "  Classic menu reverted."
        Write-Host "  Restarting Explorer..."
        Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
        Start-Sleep -Milliseconds 800
        Start-Process explorer
    }
}

# ---------------------------------------------------------------------------
# Remove installed files
# ---------------------------------------------------------------------------
$installDir = Join-Path $env:LOCALAPPDATA "ArchiveRev"
if (Test-Path $installDir) {
    # Keep the Backups subfolder - user may want their registry backups
    Get-ChildItem $installDir -File | Remove-Item -Force
    Write-Host "  Removed scripts from: $installDir"
    Write-Host "  (Registry backups preserved in: $installDir\Backups)"
}

[Microsoft.VisualBasic.Interaction]::MsgBox(
    "ArchiveRev uninstalled.`n`nRegistry backups kept at:`n$backupDir",
    64, "ArchiveRev Uninstaller"
) | Out-Null

Write-Host "`nUninstall complete."