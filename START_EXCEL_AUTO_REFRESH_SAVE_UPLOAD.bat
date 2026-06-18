@echo off
setlocal
cd /d "%~dp0"

echo.
echo ================================================
echo  EXCEL AUTO REFRESH + SAVE + WEB UPLOAD
echo ================================================
echo.
echo This will open Excel and every 60 seconds:
echo  1. Refresh Excel queries/connections
echo  2. Save the workbook
echo  3. Export OptionData to data\nifty-options.dat
echo  4. Upload DAT to GitHub Pages
echo.
echo Keep Excel and this black window open.
echo.

if not exist "nifty_option_data_cache.xlsx" (
  echo Missing Excel file: nifty_option_data_cache.xlsx
  pause
  exit /b 1
)

if not exist "excel_auto_refresh_save_upload_1min.ps1" (
  echo Missing script: excel_auto_refresh_save_upload_1min.ps1
  pause
  exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0excel_auto_refresh_save_upload_1min.ps1"

echo.
echo Auto refresh stopped.
pause
