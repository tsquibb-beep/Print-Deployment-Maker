# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

# Print Deployment Maker

## Git Workflow
- After every meaningful change, commit with a clean descriptive message and push to GitHub
- Commit message format: imperative mood, subject line under 72 characters
- Remote: `origin` → `git@github.com:tsquibb-beep/Print-Deployment-Maker.git` (SSH — not HTTPS)
- Git identity set locally: `user.name = Tom Squibb`, `user.email = tsquibb@gmail.com`
- SSH via Windows OpenSSH: `core.sshCommand = /mnt/c/Windows/System32/OpenSSH/ssh.exe`
- Always push immediately after committing — no unpushed local commits

## Versioning
- Version is stored in `version.txt` (project root) — single source of truth
- Follows SemVer: `MAJOR.MINOR.PATCH`
  - MAJOR: breaking changes / major rewrites
  - MINOR: new user-visible features
  - PATCH: bug fixes, UI polish, under-the-hood improvements
- Bump `version.txt` as part of the same commit as the change, every session
- To release: edit `version.txt`, commit, then `git tag v0.x.x && git push --tags`

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

A portable WPF PowerShell tool that generates Intune printer deployment packages. The user browses a driver `.inf` file, selects a driver model from a scrollable list, enters a deployment name, adds print queues (Name + IP), then clicks an action button to generate a `Packages\<name>\` folder containing `deploy.ps1`, `detect.ps1`, and `deployment-info.txt` — optionally packaged to `.intunewin` via IntuneWinAppUtil.exe.

**User:** Tom Squibb (tsquibb@gmail.com, GitHub: tsquibb-beep)

---

## Folder Structure

```
Print Deployment Maker\
├── Start.cmd                  ← double-click launcher
├── Start.ps1                  ← entry point; reads version.txt, calls Show-MainWindow
├── version.txt                ← SemVer single source of truth (currently 0.2.0)
├── IntuneWinAppUtil.exe       ← gitignored; must be present to use packaging buttons
├── NJK-Printer\               ← gitignored reference deployment; NEVER modify
├── Packages\                  ← gitignored runtime output
└── src\UI\
    └── MainWindow.ps1         ← all XAML + all logic (single file)
```

---

## Architecture

### Generated output per deployment type

| Button | Output folder | deploy.ps1 | detect.ps1 |
|---|---|---|---|
| Create Only | `Packages\<Name>\` | Copies driver + installs via prndrvr.vbs + adds queues | Checks printer names |
| Create and Package | `Packages\<Name>\` | Same as above + packages to .intunewin | Checks printer names |
| Package Driver Only | `Packages\<Name>-Driver\` | Copies driver + installs via prndrvr.vbs | Checks driver name |
| Package Print Queue Only | `Packages\<Name>-QueueOnly\` | Adds queues only (no driver copy) | Checks printer names |

All output folders also contain `deployment-info.txt` with driver, queues, and Intune commands.

### Generated Intune commands (all types)
- **Install:** `powershell.exe -ExecutionPolicy Bypass -File ".\deploy.ps1" -Action Install`
- **Uninstall:** `powershell.exe -ExecutionPolicy Bypass -File ".\deploy.ps1" -Action Uninstall`

### Driver installation method
```
cscript C:\Windows\System32\Printing_Admin_Scripts\en-US\prndrvr.vbs
    -a -m "<DriverName>" -h "<DriverStorePath>\" -i "<DriverStorePath>\<InfFile>"
```
Driver files are copied to `C:\ProgramData\AutoPilotConfig\Printers\<DriverFolderName>\`.

### INF parsing
`Parse-InfDriverModels` reads the `[Manufacturer]` section to find platform-specific model sections (e.g. `TOSHIBA.NTamd64`), extracts quoted model names and `%StringKey%` references (resolved via `[Strings]` section), returns sorted unique friendly names — no hardware IDs shown. Displayed in a `ListBox`, not a ComboBox.

### Theme system
`Set-Theme -Dark $true/$false` swaps all `DynamicResource` brushes (`BrushWinBg`, `BrushPanelBg`, etc.) via a palette hashtable. System highlight colour keys are also replaced. Dark mode TextBox readability requires a custom `ControlTemplate` on the implicit TextBox style — the inner `Border` must use `Background="{TemplateBinding Background}"` to bypass the WPF Aero renderer.

### Reset button
Dark red `ResetBtn` in the header bar (right of `ThemeBtn`). On click: shows `MessageBox` OK/Cancel warning, then clears all form fields, driver model list, queue list, and the four `$Script:Inf*` state variables. `InfPathBox.Foreground` is restored via `$Script:UI.Window.Resources['BrushTextFaint']` so it respects the active theme.

### Script-scope state
`$Script:InfPath`, `$Script:InfSourceDir`, `$Script:DriverFolderName`, `$Script:InfFileName`, `$Script:ScriptRoot`, `$Script:UI` (control hashtable), `$Script:IsDarkMode`

### Critical: ScriptRoot must be passed as a parameter
`$PSScriptRoot` inside `MainWindow.ps1` resolves to `src\UI\`, not the project root. **Always** pass it from `Start.ps1`:
```powershell
# Start.ps1
Show-MainWindow -AppVersion $appVersion -ScriptRoot $PSScriptRoot

# MainWindow.ps1 — Show-MainWindow
param([string]$AppVersion = '', [string]$ScriptRoot = '')
$Script:ScriptRoot = if ([string]::IsNullOrWhiteSpace($ScriptRoot)) {
    Split-Path (Split-Path $PSScriptRoot)   # fallback only
} else { $ScriptRoot }
```
Output folders (`Packages\`) and `IntuneWinAppUtil.exe` are located relative to `$Script:ScriptRoot`.

### ListView queue items use PSCustomObject
Queue items are stored as `[PSCustomObject]@{ Name = $name; IP = $ip }` — **not** `ListViewItem`. Using `ListViewItem` directly breaks `DisplayMemberBinding` because WPF treats it as its own container and the DataContext is not set. The XAML columns use `GridViewColumn.CellTemplate` / `DataTemplate` / `{Binding Name}` with explicit `Foreground="{DynamicResource BrushTextBody}"`. All PowerShell code reads `.Name` / `.IP`, never `.Content` / `.Tag`.

### Layout
Outer Grid (3 rows): header (Auto) | main area (*) | collapsible log (Auto).
Inner Grid (2 rows): scrollable shared form (*) | tab strip (Auto, MinHeight=160).
The tab strip is always visible at the bottom; the form scrolls if the window is too short.
Log pane starts **collapsed** — user clicks `▸ Log` to expand.
Header right column is a `StackPanel` containing `ResetBtn` then `ThemeBtn`.

### Per-queue print settings (staging printer)
Each queue can carry captured driver defaults (duplex, color/mono, paper, etc.).
Flow: pick the driver → **Install staging printer & open settings** creates a
throwaway local queue `PDM-Staging-<DriverFolder>` on the built-in `FILE:` port
(installing the driver locally via `pnputil /add-driver /install` + `Add-PrinterDriver`
first if needed) and opens its driver dialog with
`rundll32 printui.dll,PrintUIEntry /e /n <name>`. The user sets defaults, then
**Capture to selected queue** reads the queue's `UserPrintTicket` via .NET
`System.Printing` (`Get-StagingPrintTicket`) as PrintTicket XML, stamps `SettingsXml`
+ `SettingsSummary` onto the selected queue's PSCustomObject, and auto-removes the
staging printer (`Remove-StagingPrinter`).

The staging printer's name/IP are irrelevant — a PrintTicket is driver-schema only,
so it transfers to any same-driver queue. **Staging needs the app run elevated**
(install/remove printer + pnputil); package generation does not.

At package time, `Export-QueueSettingsFiles` writes each queue's XML to
`settings\queueN.xml` (UTF-8 **no BOM**, so `Set-PrintConfiguration` doesn't choke on
a BOM) and stamps a transient `.SettingsFile` relative path. `ConvertTo-PrinterArrayBlock`
emits `SettingsFile` into the deploy `$Printers` block. The Full and Queue-Only
`deploy.ps1` templates, after `Add-Printer`, run
`Set-PrintConfiguration -PrinterName $p.Name -PrintTicketXml (Get-Content ... -Raw)`
when a settings file is present. The XML lives in separate files — never inline in
`deploy.ps1` — so the deploy script stays plain ASCII (encoding rule).

### Generated detect.ps1 — printer detection pattern
The printer detect template uses an explicit `foreach` loop with a `$missingPrinters = @()` accumulator, matching the proven NJK-Printer reference deployment. Do **not** rewrite it as a `Where-Object` pipeline — `$null.Count` behaviour is unreliable across Intune's execution contexts. The template lives in `New-PrinterDetectScript` and produces `Write-Host` output for Intune log visibility.

### Encoding rule for generated scripts
All string literals written into generated `deploy.ps1` / `detect.ps1` content must use plain ASCII. Em dashes, curly quotes, and other Unicode characters in here-string templates cause parse failures on target devices where PowerShell reads the file as Windows-1252 instead of UTF-8.

---

## Current Status

**v0.2.0**
- Per-queue print settings: install a local staging printer from the selected driver,
  set defaults in the driver dialog, capture them per queue, and apply on target via
  `Set-PrintConfiguration` after `Add-Printer` (PrintTicket XML in `settings\queueN.xml`)
- WPF UI with 4 action buttons across themed tabs
- INF driver model parser (ListBox, scrollable)
- Full / driver-only / queue-only script generators
- `deployment-info.txt` written to every output folder
- Dark/light theme toggle (dark mode TextBox readable via custom ControlTemplate)
- Collapsible log pane (starts collapsed, no timestamps)
- IntuneWinAppUtil.exe packaging integration
- ScriptRoot bug fixed (output goes to project root `Packages\`, not `src\UI\Packages\`)
- ListView queue items visible in both themes
- Reset button in header with OK/Cancel confirmation
- deploy.ps1 encoding fix (em dash replaced with hyphen)
- detect.ps1 template matches NJK-Printer reference pattern (foreach + @() accumulator)
