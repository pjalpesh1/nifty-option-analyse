# NIFTY Option Analyse Pro

Static dashboard for NIFTY option-chain scanning using NSE option-chain data.

## Files

- `index.html` - GitHub Pages entry file.
- `nifty_option_analyse_pro.html` - same dashboard with the descriptive filename.

## NSE Site Data

Easiest free updater:

- `START_NSE_SITE_DATA.bat` - reads NSE website option-chain data, writes `data/nifty-options.dat`, and uploads it every minute when data changes.
- `run_nse_site_data_1min.ps1` - updater used by the BAT file.
- `READ_ME_NSE_SITE_DATA.txt` - quick instructions.

The current NSE website flow uses the NSE option-chain page plus:

`https://www.nseindia.com/api/option-chain-v3?type=Indices&symbol=NIFTY&expiry=...`

The website default feed is the raw GitHub DAT URL because it updates faster than GitHub Pages file cache:

`https://raw.githubusercontent.com/pjalpesh1/nifty-option-analyse/main/data/nifty-options.dat`

The parser supports official NSE option-chain JSON plus flat JSON/CSV/DAT rows with expiry, strike, CE/PE type, OI, change OI, LTP, IV, Greeks, and volume. DAT files can be comma, pipe (`|`), tab, or semicolon delimited.

Because GitHub Pages is static, direct browser fetch from NSE can be blocked by NSE cookies/CORS/session protection. If direct Load URL fails, use a small HTTPS NSE proxy that fetches the same NSE JSON server-side and enables CORS for this page, or host a generated `.dat` file on an HTTPS URL.

Auto update works with a hosted `.dat` file URL because the browser can re-fetch that URL each interval. A local file selected from the phone/PC is a one-time import and cannot be silently re-read after it changes.

Auto Start refreshes during NSE regular market time, Monday-Friday 09:15 to 15:30 IST, excluding the loaded NSE 2026 trading holidays. Outside market hours or holidays, the scanner pauses and resumes at the next open.

The browser remembers the last NSE feed URL, Auto Start preference, and last successful option-chain snapshot locally, so reopening the site can still show the last saved data while waiting for the next live refresh.

## GitHub Auto Update

The repository includes a GitHub Actions workflow:

- `.github/workflows/nse-data-autoupdate.yml`

It tries to fetch NSE option-chain data and update `data/nifty-options.dat` automatically every 5 minutes during the NSE market window. GitHub scheduled workflows cannot run faster than every 5 minutes, and GitHub can delay scheduled jobs during busy periods. NSE may also block requests from GitHub cloud IPs sometimes. For the fastest and most reliable 1-minute feed, keep using `START_NSE_SITE_DATA.bat` on the laptop.

## Excel DAT Uploader

Files:

- `nifty_option_data_cache.xlsx` - Excel template. Put live rows in the `OptionData` sheet.
- `run_excel_dat_upload_1min.ps1` - Windows uploader. Exports Excel to `data/nifty-options.dat` and uploads it to GitHub Pages.
- `START_HERE_AUTO_UPLOAD.bat` - easiest start file. Double-click this.
- `TEST_EXCEL_EXPORT_ONLY.bat` - one-time local export test. No upload.
- `START_EXCEL_AUTO_REFRESH_SAVE_UPLOAD.bat` - opens Excel, refreshes queries/connections, saves, exports DAT, and uploads every 60 seconds.

Simple use:

1. Double-click `START_HERE_AUTO_UPLOAD.bat`.
2. Excel opens.
3. Put rows in `OptionData`.
4. Save Excel.
5. Keep the black uploader window open.
6. Open the website and press `Load URL`, then `Auto Start`.

GitHub login is needed only once. The start file checks it and opens login if required. You can also set `GITHUB_TOKEN` instead of using GitHub CLI. The script uploads only when the DAT content changes.

For automatic Excel refresh/save on the laptop, use `START_EXCEL_AUTO_REFRESH_SAVE_UPLOAD.bat`. Excel must already have data in `OptionData`, or a Power Query/connection that fills `OptionData`.

If `OptionData` has 0 valid rows, upload is skipped so the website keeps the previous DAT data. For a quick test, copy rows from the `Sample` sheet into `OptionData`, then save Excel.

## DhanHQ Direct Data

For DhanHQ direct option-chain data, use the local private launcher `START_DHAN_LIVE_DATA.bat`. It asks for the Dhan access token on the PC, fetches NIFTY option chain every 60 seconds, converts it to `data/nifty-options.dat`, and uploads only the DAT file to GitHub Pages.

Do not put the Dhan access token inside the website.
