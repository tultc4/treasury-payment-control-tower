@echo off
setlocal EnableExtensions
cd /d "%~dp0"

set "PY_CMD="
where py >nul 2>nul && set "PY_CMD=py -3"
if not defined PY_CMD where python >nul 2>nul && set "PY_CMD=python"

if not defined PY_CMD (
  echo.
  echo ERROR: Python was not found on this computer.
  echo Ask IT to install Python 3.11+ or build the standalone EXE on another machine.
  echo The same Run_AP_Parser.cmd will automatically use AP_Payment_List_Parser.exe when present.
  echo.
  pause
  exit /b 1
)

%PY_CMD% -m pip install --user -r requirements.txt
if errorlevel 1 (
  echo.
  echo Installation failed. Please send the screen output to IT.
  pause
  exit /b 1
)

echo.
echo Requirements installed successfully.
pause
