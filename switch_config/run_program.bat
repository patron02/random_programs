@echo off
setlocal

rem Check if openpyxl is installed
pip show openpyxl > nul 2>&1
if errorlevel 1 (
    echo Installing openpyxl...
    pip install openpyxl
) 
powershell.exe -Command "Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass"

powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File "%~dp0switch_config_gui.ps1"

pause

endlocal
















