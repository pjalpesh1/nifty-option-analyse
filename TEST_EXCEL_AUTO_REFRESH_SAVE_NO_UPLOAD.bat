@echo off
setlocal
cd /d "%~dp0"

echo.
echo ================================================
echo  TEST EXCEL AUTO REFRESH + SAVE - NO UPLOAD
echo ================================================
echo.
echo This test opens Excel and exports local DAT only.
echo Close the black window when you want to stop.
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0excel_auto_refresh_save_upload_1min.ps1" -NoUpload

echo.
echo Test stopped.
pause
