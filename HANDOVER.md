# Session Handover — Print Deployment Maker

**Date:** 2026-06-10
**Version:** 0.3.0
**Repo:** git@github.com:tsquibb-beep/Print-Deployment-Maker.git (SSH)
**Working tree:** clean, all pushed.

---

## Starting prompt for next session

> I want to keep working on my Print Deployment Maker project (in the Projects
> directory). Read CLAUDE.md and HANDOVER.md in the project folder first, then we'll
> pick up from the suggested next steps. The app is at v0.3.0 and fully working —
> per-queue print settings (PrintTicket + DEVMODE capture) are field-verified on a
> real Toshiba Intune deployment.

---

## Where the project stands

Portable WPF PowerShell tool (`Start.cmd` → `Start.ps1` → `src\UI\MainWindow.ps1`,
single file) that generates Intune printer deployment packages. Browse a driver
`.inf`, pick a model, name the deployment, add print queues, optionally capture
per-queue driver defaults, then click one of four action buttons to produce
`deploy.ps1` / `detect.ps1` / `deployment-info.txt`, optionally packaged to
`.intunewin`.

**v0.3.0 is the current, working state. Everything below is verified working,
including on real deployments:**
- 4 action buttons (Create / Create+Package / Driver Only / Queue Only)
- INF model parser, dark/light theme, collapsible log, Reset button
- **Per-queue print settings** via a throwaway local "staging printer":
  - **PrintTicket** capture (default) — duplex, color, paper, etc.
  - **DEVMODE** capture (checkbox) — full driver DEVMODE incl. Toshiba
    Private/Hold/Scheduled print, applied on target via `Default DevMode` registry
    write + spooler restart. **Field-verified on a real Toshiba deployment.**
  - "Set" column shows ✓ for queues with captured settings; per-queue independent.

See CLAUDE.md → "Per-queue print settings (staging printer)" for the full mechanism,
data model (`SettingsBlob`/`SettingsKind`), and gotchas.

---

## How to develop & verify (important — testing from WSL)

The repo lives on a Windows path but Claude runs in WSL. The app itself runs on
Windows. Useful patterns proven this session:
- **Parse-check** a change without running the UI:
  `powershell.exe -Command "[System.Management.Automation.Language.Parser]::ParseFile(<win path>, [ref]$null, [ref]$e)"`
- **XAML load check**: dot-source `MainWindow.ps1`, then `XamlReader.Load` on
  `$Script:MainXaml` and `FindName` the new controls.
- **Headless logic tests**: dot-source the file and call functions directly
  (`Get-StagingDevmode`, `Export-QueueSettingsFiles`, `ConvertTo-PrinterArrayBlock`,
  `New-FullDeployScript`) — you can build a real `System.Windows.Controls.ListView`
  in memory and add PSCustomObjects to it.
- **Both shells work**: Windows PowerShell 5.1 and pwsh 7.x both load WPF +
  `System.Printing`. `Start.cmd` prefers pwsh.
- **Staging needs elevation** (install/remove printer, prndrvr, registry read).
- Don't run `printui.dll /Ss` to read settings — it hangs on a hidden dialog for
  some drivers. Read the `Default DevMode` registry value instead.

---

## Known gotchas (carried forward)

- **ScriptRoot** — pass `-ScriptRoot $PSScriptRoot` from `Start.ps1`; never read
  `$PSScriptRoot` directly in `MainWindow.ps1` for output paths.
- **Generated-script encoding** — all literals inside `deploy.ps1`/`detect.ps1`
  here-strings must be plain ASCII. Settings data is written to separate
  `settings\queueN.xml`/`.dat` files (XML as UTF-8 **no BOM**) so the deploy script
  stays ASCII.
- **detect.ps1** — keep the `foreach` + `$missingPrinters = @()` accumulator; don't
  refactor to a `Where-Object`/`$null.Count` pipeline.
- **Dark-mode TextBoxes** need the custom `ControlTemplate`.
- **ListView items** must be `PSCustomObject`; mutating a property needs
  `QueueListView.Items.Refresh()` to show (not an ObservableCollection).
- **Local driver install for staging** uses `prndrvr.vbs -a` (the proven method),
  NOT `pnputil` + `Add-PrinterDriver` (which fails to register the model by name).
- **`System.Printing`** is the correct assembly for `LocalPrintServer`/`PrintQueue`
  (not `ReachFramework`); `PrintTicket.SaveTo` emits a BOM that must be stripped.
- **DEVMODE capture** must be set under **Printing Defaults** (global), not
  Preferences (per-user) — that's why the DEVMODE path opens `/p`, not `/e`.
- **Git** — SSH remote; `core.sshCommand = /mnt/c/Windows/System32/OpenSSH/ssh.exe`.
  Commit + push after every meaningful change; bump `version.txt` each session.
- **IntuneWinAppUtil.exe** — gitignored; must be present in project root for
  packaging buttons.

---

## Suggested next steps (none requested yet)

1. **Queue editing** — double-click a queue row to edit Name/IP (currently
   remove + re-add).
2. **Validation polish** — red border / inline hint on empty required fields before
   a create action (errors currently only appear in the log).
3. **Uninstall removes the TCP port** — generated uninstall removes the printer but
   not the port (`Remove-PrinterPort`).
4. **Settings preview / clear** — a way to view a captured ticket's key values and a
   button to clear a queue's captured settings without removing the queue.
5. **DPI / min-size** — test at 125% and 150% scaling; confirm tab strip + buttons
   stay visible.
6. **Real logo** — replace the "PDM" text placeholder `<Border>` with a 96×96 PNG.
7. **Multiple drivers per package** — currently one driver per deployment.

---

## File map

| File | Purpose |
|---|---|
| `Start.cmd` / `Start.ps1` | Launcher; reads `version.txt`, calls `Show-MainWindow -ScriptRoot $PSScriptRoot` |
| `version.txt` | SemVer — currently `0.3.0` |
| `src\UI\MainWindow.ps1` | Everything: XAML, styles, logic, event handlers, script generators |
| `NJK-Printer\` | Reference deployment (gitignored, never modify) |
| `Packages\` | Runtime output (gitignored) |
| `IntuneWinAppUtil.exe` | Packaging tool (gitignored, must be present manually) |
