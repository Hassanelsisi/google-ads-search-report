# Google Ads Search-Campaign Reporting Script

A read-only Google Ads Script that generates an 8-tab Google Sheet with
campaign KPIs, heatmaps, optimisation tips, and brand-aware search-term handling.

## Files

| File              | Purpose |
|-------------------|---------|
| **search-report.gs** | Main script (paste into Google Ads). |
| **index.html**       | Project documentation (also renders via GitHub Pages). |

## Quick Start

1. In Google Ads, open **Tools → Bulk actions → Scripts**.  
2. Paste `search-report.gs`, replace `SPREADSHEET_ID` with your blank Sheet ID.  
3. Preview → Authorise → Run once. Schedule to refresh (e.g. at 06:00, 14:00, 22:00).

## Customising Brand Detection

Edit the array near the top of the script:

```javascript
const BRAND_PATTERNS = [
  /mybrand/i,   // ASCII
  /ماي براند/i  // Arabic
];
