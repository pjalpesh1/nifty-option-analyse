@echo off
setlocal
cd /d "%~dp0"

echo.
echo =================================
echo  TEST EXCEL TO DAT - NO UPLOAD
echo =================================
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0run_excel_dat_upload_1min.ps1" -Once -NoUpload

echo.
echo Test complete.
echo Check file: data\nifty-options.dat
echo.
pause
