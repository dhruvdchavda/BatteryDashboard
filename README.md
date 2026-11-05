# Battery Dashboard

A tiny Windows tool that turns `powercfg /batteryreport` into a clean, single-page HTML dashboard: battery health, cycle count, live charge, charge history and daily usage. No install, no dependencies, nothing leaves your PC.

## Usage

Double-click `Battery Dashboard.bat`, or run:

```powershell
.\BatteryDashboard.ps1                    # generate and open the dashboard
.\BatteryDashboard.ps1 -NoOpen            # generate only
.\BatteryDashboard.ps1 -CreateShortcut    # also add a Desktop shortcut
```

Output is written to `%LOCALAPPDATA%\BatteryDashboard\dashboard.html`.

## Requirements

- Windows 10/11 with a battery (laptop/tablet)
- Windows PowerShell 5.1 (built in)

## Privacy

Everything runs locally. The dashboard embeds your device model, BIOS and OS build, so avoid sharing the generated files publicly.

## Contributing

Issues and pull requests are welcome.

## License

[MIT](LICENSE)
