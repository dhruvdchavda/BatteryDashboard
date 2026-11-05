@echo off
rem Double-click to open the Battery Dashboard. Pass -CreateShortcut to add a Desktop shortcut.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0BatteryDashboard.ps1" %*
