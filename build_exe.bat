@echo off
setlocal

REM Build standalone EXE for Sonos-Volume-Sync using PyInstaller.
REM Requires: PyInstaller installed in the local .venv.

if not exist ".venv\Scripts\python.exe" (
    echo Python virtual environment not found at .venv\
    echo Please create it and install dependencies first:
    echo   python -m venv .venv
    echo   .venv\Scripts\python.exe -m pip install -r requirements.txt
    echo   .venv\Scripts\python.exe -m pip install pyinstaller
    goto :eof
)

echo Building sonos_volume_sync.exe ...
".venv\Scripts\python.exe" -m PyInstaller ^
  --onefile ^
  --noconsole ^
  --name sonos_volume_sync ^
  --icon sonos_volume_sync.ico ^
  --add-data "sonos_volume_sync.ico;." ^
  sonos_volume_sync.py

if errorlevel 1 (
    echo PyInstaller build failed.
    goto :eof
)

echo Build completed.
echo The EXE is in the dist\sonos_volume_sync.exe
echo Remember to copy sonos_volume_sync.config.json into the same folder as the EXE.

endlocal
