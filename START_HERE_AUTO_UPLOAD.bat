@echo off
setlocal
cd /d "%~dp0"

echo.
echo ==========================================
echo  NIFTY OPTION AUTO UPLOAD - SIMPLE START
echo ==========================================
echo.

if not exist "nifty_option_data_cache.xlsx" (
  echo Excel file missing: nifty_option_data_cache.xlsx
  echo Keep this BAT file in the same folder as the Excel file.
  pause
  exit /b 1
)

if not exist "run_excel_dat_upload_1min.ps1" (
  echo Uploader script missing: run_excel_dat_upload_1min.ps1
  pause
  exit /b 1
)

echo Opening Excel template...
start "" "%~dp0nifty_option_data_cache.xlsx"

set "GH_EXE="
where gh >nul 2>nul
if %ERRORLEVEL% EQU 0 set "GH_EXE=gh"
if not defined GH_EXE if exist "%ProgramFiles%\GitHub CLI\gh.exe" set "GH_EXE=%ProgramFiles%\GitHub CLI\gh.exe"

if not defined GH_EXE (
  echo.
  echo GitHub CLI not found.
  echo Install GitHub CLI or set GITHUB_TOKEN before running upload.
  echo The uploader can still export local DAT, but cannot publish online.
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
echo IMPORTANT:
echo 1. Put option rows in Excel sheet: OptionData
echo 2. Save Excel after every update
echo 3. Keep this black window open
echo 4. Open website and press Load URL, then Auto Start
echo.
echo Starting uploader every 60 seconds...
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0run_excel_dat_upload_1min.ps1"

echo.
echo Uploader stopped.
pause
