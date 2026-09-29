# Factor Investment Simulator (NSE)

A single-file, **100% offline** web app to compare rolling returns of NSE factor
indices — Momentum, Value, Quality, Multicap and benchmarks — for both **Lumpsum
(CAGR)** and **SIP (exact XIRR)** investing.

**Live app:** open [`index.html`](index.html) in any browser, or host it on GitHub Pages.

## What it does
- **7 indices included** (official NSE / niftyindices.com history, aligned to real
  trading days, Apr 2005 → Sep 2026):
  - Nifty 50
  - Nifty 500
  - Nifty 200 Momentum 30
  - Nifty 500 Momentum 50
  - Nifty500 Multicap Momentum Quality 50
  - Nifty 500 Quality 50
  - **Nifty 500 Value 50**
- **Price (Unadjusted)** and **TRI (Total Return)** basis toggle.
- Custom-horizon rolling returns (any number of years, incl. fractional).
- **Rolling N-Year Return** time-series (daily cadence for Lumpsum, 1st & 15th for SIP),
  **Equity Curve**, and a **2% - bin return distribution** histogram.
- **Jump to date** box + hover crosshair to inspect any window's start→end values.
- Custom blended portfolio with weight allocation across all sleeves.

No CDN, no backend, no macros — everything is embedded in the one HTML file.

## Data workbook
[`NSE_Index_Data_Price_TRI.xlsx`](NSE_Index_Data_Price_TRI.xlsx) contains the raw
history for all 7 indices in two sheets:
- **Price (Unadjusted)** — price-return index levels (Close).
- **TRI (Total Return)** — total-return index levels (dividends reinvested).

Source: official NSE Indices historical data — https://www.niftyindices.com/reports/historical-data

## Hosting on GitHub Pages
The app is `index.html` at the repo root, so once GitHub Pages is enabled
(**Settings → Pages → Source: your branch / root**) it is served directly at
`https://<username>.github.io/<repo>/`.

## Refreshing the data
The data is embedded in the HTML. To pull the latest values from NSE and regenerate
everything, run the build scripts (in `build/`, not committed to Git):

```bash
python build/refresh_data.py       # fetch latest Price + TRI, snapshot, rebuild index.html
python build/make_excel.py         # regenerate NSE_Index_Data_Price_TRI.xlsx
# offline rebuild from the last snapshot:
python build/refresh_data.py --from-cache
```

*For research / education only. Not investment advice.*
