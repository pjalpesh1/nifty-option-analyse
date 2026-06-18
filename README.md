# NIFTY Option Analyse Pro

Static dashboard for NIFTY option-chain scanning.

## Files

- `index.html` - GitHub Pages entry file.
- `nifty_option_analyse_pro.html` - same dashboard with the descriptive filename.

## Angel One Live Data

The dashboard runs as a static GitHub Pages website, but Angel One credentials must stay on a secure backend/proxy. Paste that backend HTTPS JSON/CSV feed URL in the Data tab, then click Load URL and Auto Start. The frontend can poll the backend every 1 second during NSE regular market time, Monday-Friday 09:15 to 15:30 IST, excluding the loaded NSE 2026 trading holidays. Outside market hours or holidays, the scanner pauses and resumes at the next open.

The browser remembers the last feed URL, Auto Start preference, and last successful option-chain snapshot locally, so reopening the site can still show the last saved data while waiting for the next live refresh.

If the dashboard shows no data, the browser has no saved Angel One feed/cached snapshot yet, or the saved feed URL is unreachable/CORS-blocked. Open the Data tab, load the backend feed URL, then use Auto Start.

Do not store Angel One API key, secret, JWT, feed token, PIN, password, or TOTP in this static page.
