# workstation-debloater

A menu-driven .NET 10 console utility that automates performance fixes for Windows dev workstations. Built for Dell Precision laptops but applicable to any bloated Windows developer machine.

## What it does

Run it as Administrator, pick fixes from the menu (or run them all), and confirm each action before it executes. When you quit, it writes a `restore.ps1` script that undoes everything.

```
╔══════════════════════════════════════════════════╗
║        PerformanceFixer  v1.0                    ║
║        Dell Precision 7680 Optimizer             ║
╚══════════════════════════════════════════════════╝

────────────── MENU ──────────────
[0] Run ALL fixes (with confirmation per category)
[1] Show system snapshot (before/after)
[2] Kill orphaned dotnet.exe & VBCSCompiler processes
[3] Switch power plan to High Performance
[4] Stop & disable Dell bloatware services
[5] Stop & disable redundant remote access services
[6] Disable unnecessary scheduled tasks
[7] Kill resource-heavy non-essential processes
[8] Disable unnecessary startup entries
[9] Add Windows Defender exclusions for dev paths
[Q] Quit
```

## Fix categories

| # | Category | What it does |
|---|----------|-------------|
| 1 | **System Snapshot** | Process count, total RAM, top 10 by RAM/CPU, active power plan, running service count |
| 2 | **Orphaned Processes** | Kills `dotnet.exe` not parented by Visual Studio and `VBCSCompiler.exe` — reports RAM freed |
| 3 | **Power Plan** | Switches to High Performance (creates the plan if missing) |
| 4 | **Dell Bloatware** | Stops and disables DellClientManagementService (ServiceShell), DellTechHub, SupportAssist, DellPairService, DellTrustedDevice |
| 5 | **Remote Access** | Stops and disables Splashtop, TeamViewer, SAAZ (Datto) agents, ITSPlatform — keeps ScreenConnect |
| 6 | **Scheduled Tasks** | Disables MSIAfterburner, Office Background Push Maintenance, VS UpdateConfiguration_* tasks |
| 7 | **Heavy Processes** | Kills PhoneExperienceHost, LM Studio, SDXHelper, WhatsApp |
| 8 | **Startup Entries** | Removes LM Studio, GoogleChromeAutoLaunch_*, Logi Tune from registry Run keys |
| 9 | **Defender Exclusions** | Adds path exclusions for dotnet, Visual Studio, .nuget, and source directories |

## Safety

- Every destructive action prompts **[Y/n]** before executing
- Previous service start types are recorded before changes
- Color-coded output: green = success, yellow = warning, red = error
- Generates `restore.ps1` on exit to undo all changes

## Requirements

- Windows 10/11
- [.NET 10 SDK](https://dotnet.microsoft.com/download/dotnet/10.0)
- **Administrator privileges** (the app manifest requests elevation automatically)

## Usage

```powershell
git clone https://github.com/DavidVeksler/workstation-debloater.git
cd workstation-debloater
dotnet run
```

Or build a standalone executable:

```powershell
dotnet publish -c Release -r win-x64 --self-contained
# Output: bin\Release\net10.0-windows\win-x64\publish\PerformanceFixer.exe
```

## Restoring changes

If you need to undo everything:

```powershell
# Run as Administrator
.\restore.ps1
```

The restore script is generated each time you run the tool and includes only the changes you actually applied.

## License

MIT
