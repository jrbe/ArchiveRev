# ArchiveRev

> Right-click any file or folder -> **Archive Rev** -> **Archive Rev...** or **Quick Zip** -> done.

A Windows context menu tool for anyone who snapshots work-in-progress files with revision suffixes before making big changes. Works on any file type: CAD assemblies, PCB layouts, source code, video projects, Photoshop files, documents, raw images, whatever.

---

## What it does

Right-clicking any file or folder shows an **Archive Rev** submenu with two options:

- **Archive Rev...** -- opens a dialog to set the zip filename, revision suffix, and optional date. Live preview shows the exact output filename before you commit.
- **Quick Zip** -- no dialog, instantly creates `filename_YYYY-MM-DD.zip` in the Archive folder.

Both options:
- Drop the zip into an `Archive/` subfolder (created automatically if it doesn't exist)
- Use **7-Zip's Volume Shadow Copy (`-ssw`)** so files open in SolidWorks, AutoCAD, KiCad, or any other application are captured correctly without closing them first
- Fall back to PowerShell `Compress-Archive` if 7-Zip is not installed (fallback cannot read locked files)
- Support multi-file selection -- select any number of files, right-click one, get one dialog, one zip

### Archive naming

| You right-click | Revision entered | Date checkbox | Result |
|-----------------|-----------------|---------------|--------|
| `MyPart.sldprt` | `REV_K` | checked | `Archive\MyPart_REV_K_2026-04-07.zip` |
| `MyPart.sldprt` | `REV_K` | unchecked | `Archive\MyPart_REV_K.zip` |
| `MyPart.sldprt` | *(blank)* | checked | `Archive\MyPart_2026-04-07.zip` |
| Quick Zip on any file | -- | -- | `Archive\MyPart_2026-04-07.zip` |
| `my-project\` folder | `v1.4` | checked | `my-project\Archive\my-project_v1.4_2026-04-07.zip` |

The `Archive/` folder is created automatically if it doesn't exist.

---

## Use cases

- **CAD / mechanical design** -- snapshot SolidWorks, Fusion 360, or FreeCAD files while they are open
- **PCB design** -- archive KiCad or Altium projects before a major layout change
- **Software development** -- quick project snapshot before a risky refactor (complement to git, not a replacement)
- **Video / audio production** -- checkpoint a Premiere, DaVinci, or Reaper project before restructuring
- **Graphic design** -- version Photoshop, Illustrator, or Inkscape files without a full Save As dance
- **Documents / spreadsheets** -- archive a report or model before a major revision

---

## Requirements

- Windows 10 or 11
- PowerShell 5.1+ (built into Windows -- nothing to install)
- [7-Zip](https://www.7-zip.org) *(strongly recommended -- required for open/locked file support)*

---

## First time on a new Windows machine -- allow PowerShell scripts

Windows blocks unsigned scripts by default. You only need to do this once per machine.

**1. Open PowerShell as Administrator**

Start menu -> search `powershell` -> right-click -> *Run as administrator*

**2. Set the execution policy for your user account**

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

`RemoteSigned` allows locally-created and explicitly unblocked scripts to run. It does not weaken system-wide security.

**3. Unblock the downloaded files**

Windows marks files downloaded from the internet with a hidden flag that blocks execution even after setting the policy. Unblock all scripts in one step:

```powershell
Get-ChildItem "C:\path\to\ArchiveRev\*.ps1" | Unblock-File
```

Or right-click each `.ps1` file -> **Properties** -> check **Unblock** -> OK.

> Skipping step 3 is the most common reason scripts still fail after step 2.

---

## Install

1. Download or clone this repo
2. Unblock the scripts (see above if on a fresh machine)
3. Open PowerShell in the repo folder *(no admin rights needed for the installer)*
4. Run:

```powershell
.\Install-ArchiveRev.ps1
```

The installer will:
- Copy scripts to `%LOCALAPPDATA%\ArchiveRev\`
- Register the **Archive Rev** submenu for all files and folders (current user only)
- Back up the affected registry keys before making any changes
- Ask if you want to enable the classic Windows 10-style right-click menu (recommended)

> **Windows 11 note:** By default, custom context menu items are hidden behind "Show more options."
> The installer will offer to fix this. You can also run `Toggle-ClassicMenu.ps1` any time to switch.

---

## Files included

| File | Purpose |
|------|---------|
| `ArchiveRev.ps1` | Main script -- dialog, multi-file queue, 7-Zip compression |
| `ArchiveRev_Launch.vbs` | Launcher -- runs PowerShell without a console window |
| `Install-ArchiveRev.ps1` | Installs context menu entries, backs up registry |
| `Uninstall-ArchiveRev.ps1` | Removes context menu entries, preserves backups |
| `Toggle-ClassicMenu.ps1` | Toggles Windows 11 classic right-click menu on/off |
| `Toggle-ScriptExecution.ps1` | Toggles PowerShell execution policy on/off |

---

## Toggle the classic right-click menu independently

```powershell
.\Toggle-ClassicMenu.ps1
```

Run it again to switch back. Explorer restarts automatically.

---

## Uninstall

```powershell
.\Uninstall-ArchiveRev.ps1
```

Registry backups are preserved even after uninstall.

---

## Registry and safety notes

> **Back up your registry before running any script that modifies it.**

This installer does this automatically -- a timestamped `.reg` export is saved to
`%LOCALAPPDATA%\ArchiveRev\Backups\` before any changes are made.

To restore manually from a backup:
```
reg import "path\to\backup.reg"
```

**All registry changes are scoped to `HKCU` (current user) only.** No system-wide or admin-level keys are touched.

**Note on the `*\shell` registry key:** Windows uses `*` as a literal registry key name under `HKCU\Software\Classes` to target all file types. This installer uses the .NET `Registry` API directly to write this key -- both PowerShell's registry provider and `reg.exe` incorrectly expand `*` as a wildcard on this path, causing the installer to hang. The .NET approach completes in under a second.

---

## Why the locked-file problem exists on Windows 11

Many applications hold exclusive write locks on open files. Windows 11 tightened enforcement of these locks, breaking the Windows 10 behavior that allowed right-clicking and compressing a file while it was open.

**7-Zip's `-ssw` flag** uses the **Volume Shadow Copy Service (VSS)** -- the same mechanism Windows backup tools use -- to create a read-only point-in-time snapshot before reading the file. This bypasses the lock cleanly without interrupting the application.

---

## Disclaimer

**Use at your own risk.** This software modifies the Windows registry. While all changes are reversible and scoped to the current user, the authors accept no liability for any issues arising from use of this tool. Always maintain your own backups of important data.

---

## Contributing

Issues and PRs welcome. Tested on Windows 11 24H2 with SolidWorks 2025.

---

## License

MIT
