EXCEL AUTO REFRESH + SAVE

Use:
Double-click START_EXCEL_AUTO_REFRESH_SAVE_UPLOAD.bat

What it does every 60 seconds:
1. Opens Excel workbook
2. Refreshes Excel queries/connections
3. Saves workbook
4. Exports OptionData sheet to data/nifty-options.dat
5. Uploads DAT to website

Important:
- Excel must have data in OptionData, or a Power Query/connection that fills OptionData.
- If OptionData is blank, website will show blank data.
- Save is automatic every 60 seconds while the black window is open.
- Keep Excel and black window open.

Test without upload:
Double-click TEST_EXCEL_AUTO_REFRESH_SAVE_NO_UPLOAD.bat
