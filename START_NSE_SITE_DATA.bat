@echo off
setlocal
cd /d "%~dp0"

echo.
echo ==========================================
echo  NIFTY OPTION NSE SITE DATA - AUTO START
echo ==========================================
echo.

if not exist "run_nse_site_data_1min.ps1" (
  echo NSE updater script missing: run_nse_site_data_1min.ps1
  echo Keep this BAT file in the same folder as the script.
  pause
  exit /b 1
)

set "GH_EXE="
where gh >nul 2>nul
if %ERRORLEVEL% EQU 0 set "GH_EXE=gh"
if not defined GH_EXE if exist "%ProgramFiles%\GitHub CLI\gh.exe" set "GH_EXE=%ProgramFiles%\GitHub CLI\gh.exe"

if not defined GH_EXE (
  echo.
  echo GitHub CLI not found.
  echo The updater can save local data, but cannot publish to the website.
  echo Install GitHub CLI or set GITHUB_TOKEN if you need online upload.
  echo.
) else (
  "%GH_EXE%" auth status >nul 2>nul
  if ERRORLEVEL 1 (
    echo.
    echo GitHub login needed one time. Follow the login steps.
    "%GH_EXE%" auth login
  )
)

echo.
echo This uses NSE website data, no Angel One and no Dhan token.
echo Keep this black window open.
echo Website feed URL: https://raw.githubusercontent.com/pjalpesh1/nifty-option-analyse/main/data/nifty-options.dat
echo Update cycle: every 60 seconds
echo.
echo Starting NSE updater...
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0run_nse_site_data_1min.ps1" -IntervalSeconds 60 -ExpiryCount 2

echo.
echo NSE updater stopped.
pause
