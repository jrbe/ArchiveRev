' ArchiveRev_Launch.vbs v1.0.0
' Launches ArchiveRev.ps1 without a visible console window.
' Uses WScript.Shell.Run (window style 0) which bypasses Windows Terminal
' and works correctly even when Windows Terminal is the default console host.
'
' Arguments:
'   0 : Full path to ArchiveRev.ps1
'   1 : Target path (%1 or %V from Explorer)
'   2 : Mode  (File | Folder | FolderBackground)
'   3 : (optional)  Quick  -- skips dialog, instant zip

If WScript.Arguments.Count < 3 Then WScript.Quit 1

Dim oShell, scriptPath, targetPath, mode, quickFlag, cmd
Set oShell = CreateObject("WScript.Shell")

scriptPath = WScript.Arguments(0)
targetPath = WScript.Arguments(1)
mode       = WScript.Arguments(2)
quickFlag  = ""
If WScript.Arguments.Count > 3 Then
    If WScript.Arguments(3) = "Quick" Then quickFlag = " -Quick"
End If

cmd = "powershell.exe -NoProfile -Sta -ExecutionPolicy Bypass" & _
      " -File """ & scriptPath & """" & _
      " -TargetPath """ & targetPath & """" & _
      " -Mode " & mode & quickFlag

' Window style 1 = normal. PowerShell hides its own console window immediately.
' Style 0 (hidden) prevents WinForms from rendering - the console must exist briefly.
oShell.Run cmd, 1, False