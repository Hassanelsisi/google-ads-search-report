# Google Ads Search-Campaign Reporting Script

One script, one Sheet, eight data-rich tabs.  
Paste the code into Google Ads → Bulk actions → Scripts, run it, and a
fully-formatted workbook appears in Drive. Perfect for weekly reviews
or rapid client audits.

| Tab | Highlights |
|-----|------------|
| **README** | Purpose, colour legend, disclaimer |
| **CONFIG** | Edit thresholds without code |
| **Overview** | Clicks · CTR · CPC · CPA · ROAS – currency auto-detected |
| **Heatmap** | Day × Hour grids (Clicks · CTR · CPC · Conversions · CPA) |
| **Strategy** | Rule tips, device CPA gaps, live Google-Ads recommendations |
| **Ads** | Enabled ads with colour-coded advice (keep / test / pause) |
| **Campaign-KW** | All keywords + Quality Score, Top 10 / Bottom 10, actions |
| **Search Terms** | Brand-aware Add / Review / Exclude |

---

## Quick Start

1. Open **Google Ads → Tools → Bulk actions → Scripts**.  
2. Paste `search-report.gs`.  
3. Replace `SPREADSHEET_ID` with an empty Google Sheet ID.  
4. **Preview → Authorise → Run** once.  
5. Schedule the script (e.g. 06:00, 14:00, 22:00) to refresh automatically.

---

## Colour legend

| Colour | Meaning |
|--------|---------|
| 🟩 `#c8e6c9` | Good / Add / Keep / Highest CTR / Lowest CPC |
| 🟥 `#ffcdd2` | Exclude / Pause / Lowest CTR / Highest CPC |
| 🟧 `#ffe0b2` | Review |
| 🟨 `#fff9c4` | High CPC (above multiplier) |
| 🟦 `#bbdefb` | Low CPC (below average) |

---

## CONFIG keys

| Key | What it controls |
|-----|------------------|
| `CLICK_THRESHOLD` | Clicks before a keyword or term may be excluded |
| `CTR_HIGH` / `CTR_LOW` | Boundaries that trigger “High CTR” / “Low CTR” tips |
| `CPC_MULTIPLIER` | High-CPC flag if Avg CPC > multiplier × account avg |
| `CPA_MULTIPLIER` | High-CPA flag if CPA > multiplier × account avg |

Change values in the Sheet’s CONFIG tab and re-run—no code edits needed.

---

## Customising brand detection

Edit the array near the top of the script:

```js
const BRAND_PATTERNS = [
  /mybrand/i,      // ASCII
  /ماي براند/i     // Arabic
];
