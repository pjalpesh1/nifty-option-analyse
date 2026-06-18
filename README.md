# NIFTY Option Analyse Pro

Static dashboard for NIFTY option-chain scanning using NSE option-chain data.

## Files

- `index.html` - GitHub Pages entry file.
- `nifty_option_analyse_pro.html` - same dashboard with the descriptive filename.

## NSE Site Data

The dashboard defaults to the NSE NIFTY option-chain JSON URL:

`https://www.nseindia.com/api/option-chain-indices?symbol=NIFTY`

The parser supports official NSE option-chain JSON plus flat JSON/CSV/DAT rows with expiry, strike, CE/PE type, OI, change OI, LTP, IV, Greeks, and volume. DAT files can be comma, pipe (`|`), tab, or semicolon delimited.

Because GitHub Pages is static, direct browser fetch from NSE can be blocked by NSE cookies/CORS/session protection. If direct Load URL fails, use a small HTTPS NSE proxy that fetches the same NSE JSON server-side and enables CORS for this page, or host a generated `.dat` file on an HTTPS URL.

Auto update works with a hosted `.dat` file URL because the browser can re-fetch that URL each interval. A local file selected from the phone/PC is a one-time import and cannot be silently re-read after it changes.

Auto Start refreshes during NSE regular market time, Monday-Friday 09:15 to 15:30 IST, excluding the loaded NSE 2026 trading holidays. Outside market hours or holidays, the scanner pauses and resumes at the next open.

The browser remembers the last NSE feed URL, Auto Start preference, and last successful option-chain snapshot locally, so reopening the site can still show the last saved data while waiting for the next live refresh.
