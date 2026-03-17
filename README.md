# PC Gaming Tweaks + Scripts

A personal collection of PowerShell scripts, registry tweaks, and Windows utilities focused on PC gaming performance, diagnostics, and system management.

---

## Repository Structure

```
PC-Gaming-Tweaks-Scripts/
├── diagnostics/        # System info + hardware diagnostic scripts
├── tweaks/             # Registry edits and performance optimizations
├── utilities/          # General-purpose Windows helper scripts
└── README.md
```

---

## Scripts

### Diagnostics

| Script | Description |
|---|---|
| `diagnostics/SystemInfo.ps1` | Exports DXDiag, ComputerInfo, and SystemInfo reports to the Desktop |

---

## Requirements

- Windows 10 or Windows 11
- PowerShell 5.1 or later (most scripts)
- Run as **Administrator** where noted

---

## Usage

Right-click any `.ps1` file and select **Run with PowerShell**, or open a PowerShell window and run:

```powershell
.\ScriptName.ps1
```

Some scripts may require bypassing the execution policy:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

---

## Output Locations

Most scripts write output files to `%USERPROFILE%\Desktop` or `%USERPROFILE%` unless otherwise noted in the script header.

---

## Notes

- Scripts are written for personal use and provided as-is.
- Always review a script before running it on your system.
- Some tweaks are system-specific. Test in a safe environment before applying broadly.

---

*Built and maintained by Wesley Manning*
