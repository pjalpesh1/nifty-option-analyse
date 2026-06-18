# NIFTY Option Analyse Pro

Static dashboard for NIFTY option-chain scanning.

## Files

- `index.html` - GitHub Pages entry file.
- `nifty_option_analyse_pro.html` - same dashboard with the descriptive filename.

## Angel One Live Data

The dashboard runs as a static GitHub Pages website, but Angel One credentials must stay on a secure backend/proxy. Paste that backend HTTPS JSON/CSV feed URL in the Data tab, then click Load URL and Auto Start. The frontend can poll the backend every 1 second.

Do not store Angel One API key, secret, JWT, feed token, PIN, password, or TOTP in this static page.
