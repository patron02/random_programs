@echo off
setlocal

powershell.exe -Command "Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass"

powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File "%~dp0iso_gui.ps1"

pause

endlocal