@echo off
rem IMPORTANT: Replace "C:\Users\VerNe\Downloads\Documents\AppPowerSwitcher" with the actual full path to your project root directory.
set "PROJECT_DIR=C:\Users\VerNe\Downloads\Documents\AppPowerSwitcher"

rem Change directory to the project root. The /d is needed if the project is on a different drive.
cd /d "%PROJECT_DIR%"

rem IMPORTANT: Replace "C:\Users\VerNe\Downloads\Documents\AppPowerSwitcher\.venv\Scripts\pythonw.exe" with the actual full path to pythonw.exe in your virtual environment or Python installation.
set "PYTHONW_PATH=%PROJECT_DIR%\.venv\Scripts\pythonw.exe"

rem Start the main application script using pythonw.exe to run silently.
rem Use start "" to ensure the Pythonw process doesn't inherit the BAT window handle, making the BAT return immediately.
rem The "" is a dummy title for the new window (though pythonw won't show a window title).
start "" "%PYTHONW_PATH%" main.py

rem Optional: Add error logging if needed (less useful for silent run).
rem If you need to debug BAT issues, remove "@echo off" and "start """.

exit /b 0