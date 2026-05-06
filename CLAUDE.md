# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

# Print Deployment Maker

## Git Workflow
- After every meaningful change, commit with a clean descriptive message and push to GitHub
- Commit message format: imperative mood, subject line under 72 characters
- Remote: `origin` → https://github.com/tsquibb-beep/Print-Deployment-Maker
- Always push immediately after committing — no unpushed local commits

## Versioning
- Version is stored in `version.txt` (project root) — single source of truth
- Follows SemVer: `MAJOR.MINOR.PATCH`
  - MAJOR: breaking changes / major rewrites
  - MINOR: new user-visible features
  - PATCH: bug fixes, UI polish, under-the-hood improvements
- To release: edit `version.txt`, commit, then `git tag v1.x.x && git push --tags`

---

## Commands

### Launch the app
```batch
Start.cmd
```
Prefers `pwsh.exe` (PS7), falls back to `powershell.exe` 5.1. No bootstrap or module install required.

### Dev loop
Edit `src\UI\MainWindow.ps1` → re-run `Start.cmd`.

---

## Project Overview

A portable WPF PowerShell tool that generates Intune printer deployment packages. The user browses a driver `.inf` file, selects a driver model, enters a deployment name, adds print queues (Name + IP), then clicks an action button to generate a `Packages\<name>\` folder containing `deploy.ps1` and `detect.ps1` — optionally packaged to `.intunewin` via IntuneWinAppUtil.exe.

**User:** Tom Squibb (tsquibb@gmail.com, GitHub: tsquibb-beep)

---

## Folder Structure

```
Print Deployment Maker\
├── Start.cmd                  ← double-click launcher
├── Start.ps1                  ← entry point; reads version.txt, calls Show-MainWindow
├── version.txt                ← SemVer single source of truth
├── IntuneWinAppUtil.exe       ← gitignored; must be present to use packaging buttons
├── NJK-Printer\               ← gitignored reference deployment; never modify
├── Packages\                  ← gitignored runtime output
└── src\UI\
    └── MainWindow.ps1         ← all XAML + all logic (single file)
```

---

## Architecture

### Generated output per deployment type

| Button | Output folder | deploy.ps1 | detect.ps1 |
|---|---|---|---|
| Create / Create and Package | `Packages\<Name>\` | Copies driver + installs via prndrvr.vbs + adds queues | Checks printer names |
| Package Driver Only | `Packages\<Name>-Driver\` | Copies driver + installs via prndrvr.vbs | Checks driver name |
| Package Print Queue Only | `Packages\<Name>-QueueOnly\` | Adds queues only (no driver copy) | Checks printer names |

### Generated Intune commands (all types)
- **Install:** `powershell.exe -ExecutionPolicy Bypass -File ".\deploy.ps1" -Action Install`
- **Uninstall:** `powershell.exe -ExecutionPolicy Bypass -File ".\deploy.ps1" -Action Uninstall`

### Driver installation method (from reference NJK-Printer deployment)
```
cscript C:\Windows\System32\Printing_Admin_Scripts\en-US\prndrvr.vbs
    -a -m "<DriverName>" -h "<DriverStorePath>\" -i "<DriverStorePath>\<InfFile>"
```
Driver files are copied to `C:\ProgramData\AutoPilotConfig\Printers\<DriverFolderName>\`.

### INF parsing
`Parse-InfDriverModels` in MainWindow.ps1 reads the `[Manufacturer]` section to find platform-specific model sections (e.g. `TOSHIBA.NTamd64`), extracts quoted model names and `%StringKey%` references (resolved via `[Strings]` section), returns sorted unique friendly names — no hardware IDs shown.

### Theme system
`Set-Theme -Dark $true/$false` swaps all `DynamicResource` brushes (`BrushWinBg`, `BrushPanelBg`, etc.) identically to the Intune Bulk Assigner pattern.

### Script-scope state
`$Script:InfPath`, `$Script:InfSourceDir`, `$Script:DriverFolderName`, `$Script:InfFileName`, `$Script:ScriptRoot`, `$Script:UI` (control hashtable), `$Script:IsDarkMode`

---

## Current Status

**v0.1.0 — initial implementation complete.**
- WPF UI with all 4 action buttons
- INF driver model parser
- Full / driver-only / queue-only script generators
- Dark/light theme toggle
- Scrollable log pane
- IntuneWinAppUtil.exe packaging integration
