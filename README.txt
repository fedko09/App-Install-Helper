MicroDeploy Toolbox – README.txt

MicroDeploy Toolbox is a unified deployment, updating, and removal utility for Windows 10 and 11.
It provides a single, clean graphical interface for installing software, updating applications, or performing deep uninstalls — all powered by Winget and Chocolatey behind the scenes.

The goal is simple: no more hunting installers, no more manual updates, and no more inconsistent uninstallers. This tool gives normal users, gamers, power users, and IT professionals an all-in-one deployment panel.

What MicroDeploy Toolbox Does
✔ Centralized Install / Update / Uninstall GUI

Select applications by checkbox, choose a backend (Winget or Chocolatey), and run an installation or removal with one click.
All actions are logged in real time inside the app.

✔ Profiles (Gamer / Everyday User / Power User / SysAdmin)

Each profile automatically filters the app list so only the relevant tools are shown.
This keeps the UI clean and prevents noise.

✔ Auto-Detection of Installed Apps

When Winget supports JSON output, the toolbox marks apps as Installed and highlights them in bold.

✔ App Runtime Management (Update / Uninstall)

When updating or uninstalling:

If the application has known background processes or services, the toolbox detects them.

For updates → prompts the user to stop them.

For uninstall → automatically kills the processes and services.

After update → attempts to relaunch stopped processes.

After uninstall → performs filesystem + registry cleanup using known vendor tokens.

✔ Automatic Backend Handling

If Winget is missing → prompts user to install/repair via Microsoft Store.

If Chocolatey is missing → offers auto-installation and elevates if needed.

If Chocolatey is installed → automatically checks for core updates once per run.

✔ Hover Tooltips

Every app has a short description visible on mouse-hover.

✔ Progress Overlay ("Working…")

Whenever the toolbox is busy (installing, scanning, updating), a dark translucent overlay appears with:

A progress indicator

Action description

Prevents UI freeze confusion

✔ JSON Profile Management

Save and load application selections as reusable .json profiles:

Per-department deployments

Golden image presets

Rapid reconfiguration for new users or new systems

✔ Logging

Every action is timestamped in the output panel:

Backend used

Operation results

Winget/Choco return codes

Cleanup actions

Errors and handling steps

Included Functions (Internal Overview)

Add-App
Creates the metadata record including Winget ID, Chocolatey package ID, processes, services, cleanup tokens, description, and profile visibility.

Apply-Profile
Shows/hides apps based on the chosen profile.

Get-SelectedApps
Pulls all visible + checked items for deployment actions.

Ensure-Winget / Ensure-Choco
Detects presence, triggers repair, auto-installs if missing.

Stop-AppRuntime / Restart-AppRuntime
Detects running services and processes, stops them for update/uninstall, and relaunches where appropriate.

Cleanup-AppAfterUninstall
Deletes leftover program folders and registry entries using vendor name tokens.

Invoke-AppAction
The core engine that drives install, update, and uninstall.
Handles backend selection, return code mapping, stopping/restarting runtimes, post-uninstall cleanup, and output logging.

Detect-InstalledApps
Uses Winget’s JSON mode (if supported) to mark apps installed.

Save-Profile / Load-Profile
Export/import selected apps and backend preferences as JSON.

Requirements

Windows 10 or Windows 11

PowerShell 5.1 or PowerShell 7+

Internet access for initial installs via Winget / Chocolatey

Administrator rights recommended for:

Deep uninstalls

Registry and program data cleanup

Chocolatey installation

Author

Developed by: Bogdan Fedko
MicroDeploy Toolbox – Automated deployment for all user types.