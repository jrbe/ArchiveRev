# Toggle-ScriptExecution.ps1 v1.0.0
# Toggles PowerShell script execution for the current user between:
#   RemoteSigned  - allows local/unblocked scripts to run (developer-friendly)
#   Restricted    - blocks all scripts (Windows default on new machines)
#
# Does NOT require admin rights (CurrentUser scope only).
# Your previous policy is saved and restored on the next toggle.
#
# USE AT YOUR OWN RISK. See README for details.

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName Microsoft.VisualBasic

# Storage key for saving the previous policy
$savedPolicyPath = "HKCU:\Software\ArchiveRev"
$savedPolicyKey  = "PreviousExecutionPolicy"

# Current effective policy for this user
$current = Get-ExecutionPolicy -Scope CurrentUser

# Read any previously saved policy
$savedPolicy = $null
if (Test-Path $savedPolicyPath) {
    $savedPolicy = (Get-ItemProperty -Path $savedPolicyPath -ErrorAction SilentlyContinue).$savedPolicyKey
}

if ($current -eq "RemoteSigned" -or $current -eq "Unrestricted" -or $current -eq "Bypass") {
    # Currently open - offer to lock it back down
    $revertTo = if ($savedPolicy) { $savedPolicy } else { "Restricted" }

    $choice = [Microsoft.VisualBasic.Interaction]::MsgBox(
        "Script execution is currently ENABLED ($current).`n`n" +
        "Restrict it back to: $revertTo`n`n" +
        "You will not be able to run .ps1 scripts until you enable it again.",
        4 + 32, "Toggle Script Execution"
    )
    if ($choice -ne 6) { exit 0 }

    Set-ExecutionPolicy -ExecutionPolicy $revertTo -Scope CurrentUser -Force
    [Microsoft.VisualBasic.Interaction]::MsgBox(
        "Script execution set to: $revertTo`n`nRun this script again to re-enable.",
        64, "Toggle Script Execution"
    ) | Out-Null

} else {
    # Currently locked - offer to enable RemoteSigned
    $choice = [Microsoft.VisualBasic.Interaction]::MsgBox(
        "Script execution is currently RESTRICTED ($current).`n`n" +
        "Enable RemoteSigned for current user?`n`n" +
        "RemoteSigned allows local and explicitly unblocked scripts to run.`n" +
        "Scripts downloaded from the internet still require unblocking first.",
        4 + 32, "Toggle Script Execution"
    )
    if ($choice -ne 6) { exit 0 }

    # Save current policy so we can restore it on the next toggle
    if (-not (Test-Path $savedPolicyPath)) {
        New-Item -Path $savedPolicyPath -Force | Out-Null
    }
    Set-ItemProperty -Path $savedPolicyPath -Name $savedPolicyKey -Value $current.ToString()

    Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
    [Microsoft.VisualBasic.Interaction]::MsgBox(
        "Script execution set to: RemoteSigned`n`n" +
        "Remember to also Unblock downloaded .ps1 files:`n" +
        "  Get-ChildItem '.\*.ps1' | Unblock-File`n`n" +
        "Run this script again to restrict back to: $current",
        64, "Toggle Script Execution"
    ) | Out-Null
}