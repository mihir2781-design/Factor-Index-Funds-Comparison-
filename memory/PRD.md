# PRD — Factor Investment Simulator, Rolling Pro (.xlsm)

## Problem statement
Deliver one hardened macro-enabled Excel workbook (`Factor_Investment_Simulator_Rolling_Pro.xlsm`) that compares Momentum, Value, Benchmark, and a custom blended factor portfolio via rolling-return analytics, with file-based NSE data import, VBA-driven engine, and a minimal dashboard.

## User choices (locked)
- Data import: **file-based** (NSE/NSE-Indices CSV exports).
- Delivery: **seeded** with realistic synthetic history (2009–2025) so it runs out of the box.
- Snapshots: **no** versioning — one clean daily table (overwrite/append).
- SIP: **exact XIRR fully in VBA** (Newton–Raphson + bisection fallback).
- Design: clean minimal professional finance theme (agent default: navy/slate + amber).

## Users
Quant investors, factor-cycle analysts, SIP-vs-Lumpsum advisors, Excel power users.

## Architecture
- Build tooling (Linux, no Excel): `pyOpenVBA` generates a valid `vbaProject.bin` from `/app/vba/*`; `XlsxWriter` assembles the workbook (sheets, seed data, named ranges, dropdowns, formatting, macro-wired buttons) and embeds the VBA project.
- Sheets: Dashboard (Sheet1) / Database_Daily / Config / Calc_Cache (hidden).
- VBA modules: modValidation, modBlendEngine, modRollingReturns, modDataImport, modDashboard, ThisWorkbook (Workbook_Open).

## Implemented (2026-06)
- Full workbook shell with named ranges, SIP/Lumpsum + 3/5/10 dropdowns, allocation total formula.
- File-based NSE import: header auto-detection (Date + TRI/value), multi-format date parsing, merge/sort/forward-fill, non-destructive when no files.
- Engine: monthly-cadence rolling window generator; Lumpsum CAGR; exact SIP XIRR; fixed-split blend (no rebalance).
- Outputs grid: Avg / Worst / Best / P>12% / P>15% / Projected Corpus for Momentum, Value, Benchmark, Blended; window count + timestamp; audit cache.
- Hard validation + explicit `INSUFFICIENT DATA`.
- Sample NSE CSV exports in `/app/sample_nse_exports/`.

## Verification
- pyOpenVBA structural validate() = clean; OOXML wiring confirmed (vbaProject rel/content-type, ThisWorkbook + Sheet1 codenames, button→macro bindings, defined names).
- Engine logic verified via faithful Python port on identical seed data: windows generate (96/84/36), metrics sane, XIRR converges, too-short range → INSUFFICIENT DATA.
- Final macro execution occurs inside Microsoft Excel (cannot run VBA in this Linux container).

## Fixes / Iterations
## Backlog / P1-P2
- 2026-06 (it.1): Fixed Momentum-underperforms bug by re-seeding synthetic data (later superseded by real data).
- 2026-06 (it.2): REAL NSE DATA + CHARTS + WEB REFRESH.
  - Discovered the live niftyindices.com TRI endpoint (POST /BackPage/getTotalReturnIndexString, cookie-primed) and index-name mapping. build/fetch_nse.py pulls full real TRI history for all 3 indices (5266 aligned daily rows, 2005-04-01 -> 2026-06-25) into build/nse_tri_real.csv; build_workbook.load_dataset() seeds Database_Daily from it (synthetic is now only a fallback). Real CAGR: Momentum 18.85% > Value 16.88% > Benchmark 14.11%.
  - Added two Dashboard charts (Equity Curve line rebased to 100; Rolling-Return Distribution column histogram) bound to a hidden Charts_Data sheet; modCharts.BuildChartData fills it on Run Simulation.
  - Added Web Refresh button + modWebRefresh (MSXML2 direct TRI download fallback, JSON scan parser) + Config web section (cfg_WebEnable/Start/End). File import remains primary.
  - Verified by testing agent: 16/16 backend checks, incl. live NSE endpoint reachable and chart/button coexistence.

## Legacy Backlog
- 2026-06 (it.3): Built self-contained OFFLINE HTML web simulator (/app/factor_simulator.html, served copy /app/frontend/public/simulator.html). 602KB single file, no CDN/backend. Embeds official NSE Price + TRI for 6 indices (Nifty 50, Nifty 500, Momentum 30, Momentum 50, Multicap MQ 50, Quality 50), aligned to 5331 trading days 2005-04-01..2026-09-28. Free-text Horizon (any years incl. fractional), TRI/Price toggle, 6-way custom blend, dependency-free canvas charts (equity curve + distribution). Engine verified via Node to match browser and the Python port. GitHub-Pages hostable (rename to index.html). Build: build/build_html.py + build/template.html.
- Optional direct web-download refresh from NSE (fragile) as fallback to file import.
- Rolling-return distribution chart / equity-curve chart on Dashboard.
- Optional daily (vs monthly) rolling cadence toggle.
- Per-index missing-file handling that still runs single-asset columns.

- 2026-06 (it.4): Added 7th index Nifty 500 Value 50 (key `v50`, color #2E86DE blue) with official Price + TRI fetched fresh from niftyindices.com; re-ran build/build_html.py (695KB, 5331 aligned days, range unchanged 2005-04-01..2026-09-28). Appears in picker/weights/table/charts/legend. Reduced Lumpsum rolling-return line width 1.8->1.1px (SIP unchanged, equity curve unchanged) — drawLine() keys off st.mode==="Lumpsum". Verified via screenshot.

- 2026-06 (it.5): Added "Jump to date" box above the Rolling Return chart (date input + Go + Latest buttons). jumpToDate() snaps crosshair to the nearest window end date, updates rollReadout, and syncs the equity chart to the same date; syncJumpInput() sets min/max/value on each run. Added build/refresh_data.py: fetches latest Price+TRI for all 7 indices from niftyindices.com, saves dated + latest.json snapshots under build/data_snapshots/, regenerates factor_simulator.html + frontend/public/simulator.html. Supports `--from-cache` for offline rebuild. Verified via screenshot (jump) and CLI (fetch + cache rebuild).
