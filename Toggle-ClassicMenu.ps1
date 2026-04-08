# Toggle-ClassicMenu.ps1 v1.0.0
# Toggles the Windows 11 right-click context menu between:
#   Classic (full menu shown immediately - Windows 10 behavior)
#   Modern  (abbreviated menu with "Show more options" - Windows 11 default)
#
# No admin rights required. Applies to current user only.
# Explorer is restarted automatically to apply the change.
#
# USE AT YOUR OWN RISK. Back up your registry before running.
# See README for backup instructions.

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName Microsoft.VisualBasic

$regPath = "HKCU:\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32"

# Detect current state
$classicActive = Test-Path $regPath

if ($classicActive) {
    $action = "DISABLE classic menu (restore Windows 11 'Show more options' behavior)"
    $confirm = "Revert to the Windows 11 abbreviated right-click menu?`n`n'Show more options' will be required again to see the full menu."
} else {
    $action = "ENABLE classic menu (show full menu immediately - Windows 10 style)"
    $confirm = "Restore the classic Windows 10-style right-click menu?`n`nThe full menu will appear immediately on right-click - no 'Show more options' needed."
}

$choice = [Microsoft.VisualBasic.Interaction]::MsgBox(
    "$confirm`n`nExplorer will restart to apply the change.",
    4 + 32,   # YesNo + Question
    "Toggle Classic Right-Click Menu"
)

if ($choice -ne 6) { exit 0 }   # 6 = Yes

if ($classicActive) {
    # Revert to Win11 modern menu
    Remove-Item -Path "HKCU:\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}" -Recurse -Force
    $msg = "Reverted to Windows 11 modern menu.`n'Show more options' is required again."
} else {
    # Enable classic full menu
    New-Item -Path $regPath -Force | Out-Null
    Set-ItemProperty -Path $regPath -Name "(default)" -Value "" -Force
    $msg = "Classic menu enabled.`nFull right-click menu will show immediately."
}

# Restart Explorer to apply
Write-Host "Restarting Explorer..."
Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
Start-Sleep -Milliseconds 800
Start-Process explorer

[Microsoft.VisualBasic.Interaction]::MsgBox($msg, 64, "Toggle Classic Right-Click Menu") | Out-Null