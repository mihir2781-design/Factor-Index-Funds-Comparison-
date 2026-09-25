# Factor Investment Simulator — Rolling Pro

A single hardened, macro-enabled Excel workbook — **`Factor_Investment_Simulator_Rolling_Pro.xlsm`** — for comparing **Momentum**, **Value**, **Benchmark**, and a **custom blended** factor portfolio using rolling-return analytics. All heavy logic (data import, rolling windows, CAGR, exact SIP XIRR, blend engine) runs in VBA.

## What you get
- `Factor_Investment_Simulator_Rolling_Pro.xlsm` — the deliverable (open in Microsoft Excel and **Enable Macros**).
- `vba/` — the full, auditable VBA source (also embedded in the workbook).
- `sample_nse_exports/` — sample NSE-Indices-style TRI CSV files to test the Refresh Data flow.
- `build_workbook.py` — reproducible build script.

## Workbook layout
- **Dashboard** — inputs, action buttons (Refresh Data / Run Simulation / Reset Inputs), status line, and the results comparison grid.
- **Database_Daily** — normalized daily TRI history (Date, Momentum, Value, Benchmark). Pre-seeded with ~17 years (2009–2025) of realistic synthetic data so it runs out of the box.
- **Config** — CSV import file paths + assumptions/notes.
- **Calc_Cache** — hidden per-window audit log (entry/maturity dates + per-asset returns).

## How to use
1. Open the workbook, **Enable Macros**.
2. *(Optional)* On **Config**, set the three CSV paths, then click **Refresh Data** to import official NSE exports (Date + TRI columns auto-detected; histories merged, sorted, forward-filled). Leave paths blank to keep the seeded data.
3. On **Dashboard**, enter: Mode (`SIP`/`Lumpsum`), Amount (₹), Horizon (`3`/`5`/`10`), Start/End dates, and Momentum% + Value% (must total **100**).
4. Click **Run Simulation**.

## Engine
- **Lumpsum → rolling CAGR.** **SIP → exact rolling XIRR** (native VBA Newton–Raphson with a bisection fallback — no Excel-function dependency).
- Rolling windows are generated at a **monthly cadence** inside the selected sample period; each window needs a full N-year horizon that ends on/before the End date.
- **Blended portfolio** applies a fixed Momentum/Value split at every contribution — no rebalancing afterward.
- Outputs per asset: Average Rolling Return, Worst-Case Floor, Best-Case Peak, P(return > 12%), P(return > 15%), and Final Projected Maturity Corpus (₹).

## Hard validation (never silently corrected)
- Mode must be exactly `SIP` or `Lumpsum`; Horizon exactly `3`/`5`/`10`; Amount > 0.
- Momentum % + Value % must equal **100** (else `INVALID ALLOCATION`).
- Start < End; a range shorter than the horizon returns **`INSUFFICIENT DATA`**, as does any sample with no complete rolling window.

## Rebuild
```bash
pip install openpyxl xlsxwriter pyOpenVBA
python3 build_workbook.py          # rebuild the .xlsm
python3 build/make_samples.py      # regenerate sample NSE CSVs
```

> Note: this is an Excel/VBA artifact — verification here was done by (a) structural validation of the embedded VBA project and OOXML wiring, and (b) a faithful Python port of the engine run against the identical seed data (window generation, CAGR, XIRR, and edge cases). Final macro execution happens inside Microsoft Excel on macro-enabled open.
