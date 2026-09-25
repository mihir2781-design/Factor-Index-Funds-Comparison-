#!/usr/bin/env python3
"""Build Factor_Investment_Simulator_Rolling_Pro.xlsm.

Pipeline:
  1. Generate a valid vbaProject.bin from the /app/vba source files (pyOpenVBA).
  2. Generate realistic synthetic daily TRI history for the three indices.
  3. Assemble the workbook (sheets, data, named ranges, dropdowns, formatting,
     macro-wired buttons) with XlsxWriter and embed the VBA project.
"""

import os
import math
import random
import zipfile
from datetime import date, timedelta

import xlsxwriter
from pyopenvba import ExcelFile, VBAModuleKind

APP = "/app"
VBA_DIR = os.path.join(APP, "vba")
BUILD_DIR = os.path.join(APP, "build")
OUT = os.path.join(APP, "Factor_Investment_Simulator_Rolling_Pro.xlsm")
os.makedirs(BUILD_DIR, exist_ok=True)


def read_src(name):
    with open(os.path.join(VBA_DIR, name), "r", encoding="utf-8") as f:
        return f.read().replace("\r\n", "\n").replace("\n", "\r\n")


def build_vba_project():
    base = os.path.join(BUILD_DIR, "_base.xlsm")
    binp = os.path.join(BUILD_DIR, "vbaProject.bin")
    for p in (base, binp):
        if os.path.exists(p):
            os.remove(p)

    wb = ExcelFile.create_new(base)
    proj = wb.vba_project()

    # Document module for the workbook.
    wb.set_module("ThisWorkbook", read_src("ThisWorkbook.cls"))

    # Remove the placeholder standard module from the template.
    try:
        proj.delete_module("Module1")
    except Exception:
        pass

    for mod in ("modValidation", "modBlendEngine", "modRollingReturns",
                "modDataImport", "modDashboard"):
        proj.add_module(mod, read_src(mod + ".bas"), kind=VBAModuleKind.standard)

    wb.save()
    wb.close()

    with zipfile.ZipFile(base) as z:
        data = z.read("xl/vbaProject.bin")
    with open(binp, "wb") as f:
        f.write(data)

    # sanity read-back
    print("VBA modules embedded:", ExcelFile(base).module_names())
    return binp


def business_days(start, end):
    d = start
    out = []
    while d <= end:
        if d.weekday() < 5:  # Mon-Fri
            out.append(d)
        d += timedelta(days=1)
    return out


def generate_series(dates):
    """Correlated geometric brownian motion for three TRI series."""
    random.seed(20240607)
    dt = 1.0 / 252.0
    specs = {  # (mu, sigma, start)
        "mom": (0.160, 0.200, 1000.0),
        "val": (0.130, 0.180, 1000.0),
        "ben": (0.115, 0.150, 1000.0),
    }
    rho = 0.70
    vals = {k: [v[2]] for k, v in specs.items()}
    for _ in range(1, len(dates)):
        common = random.gauss(0, 1)
        for k, (mu, sigma, _s) in specs.items():
            idio = random.gauss(0, 1)
            z = rho * common + math.sqrt(1 - rho * rho) * idio
            prev = vals[k][-1]
            nxt = prev * math.exp((mu - 0.5 * sigma * sigma) * dt + sigma * math.sqrt(dt) * z)
            vals[k].append(nxt)
    return vals


def build_workbook(vba_bin):
    dates = business_days(date(2009, 1, 1), date(2025, 12, 31))
    series = generate_series(dates)

    wb = xlsxwriter.Workbook(OUT, {"in_memory": True})
    wb.set_vba_name("ThisWorkbook")

    # ---- palette / formats (professional dark-slate + amber finance theme) ----
    NAVY = "#0F2A43"
    SLATE = "#1B3A57"
    STEEL = "#2E5A7D"
    AMBER = "#E0A106"
    LIGHT = "#F4F7FA"
    GREY = "#5A6B7B"

    f_title = wb.add_format({"bold": True, "font_size": 20, "font_color": "white",
                             "bg_color": NAVY, "align": "left", "valign": "vcenter",
                             "font_name": "Georgia", "indent": 1})
    f_sub = wb.add_format({"font_size": 10, "font_color": "#C9D6E2", "bg_color": NAVY,
                           "align": "left", "valign": "vcenter", "indent": 1, "italic": True})
    f_section = wb.add_format({"bold": True, "font_size": 12, "font_color": "white",
                               "bg_color": SLATE, "align": "left", "valign": "vcenter", "indent": 1})
    f_label = wb.add_format({"font_size": 11, "font_color": "#22303C", "align": "left",
                             "valign": "vcenter", "indent": 1})
    f_input = wb.add_format({"font_size": 11, "bold": True, "font_color": NAVY, "bg_color": "#FFF6DD",
                             "border": 1, "border_color": AMBER, "align": "center", "valign": "vcenter"})
    f_input_cur = wb.add_format({"font_size": 11, "bold": True, "font_color": NAVY, "bg_color": "#FFF6DD",
                                 "border": 1, "border_color": AMBER, "align": "center", "valign": "vcenter",
                                 "num_format": '\u20b9 #,##0'})
    f_input_date = wb.add_format({"font_size": 11, "bold": True, "font_color": NAVY, "bg_color": "#FFF6DD",
                                  "border": 1, "border_color": AMBER, "align": "center", "valign": "vcenter",
                                  "num_format": "dd-mmm-yyyy"})
    f_total = wb.add_format({"font_size": 11, "bold": True, "font_color": "white", "bg_color": STEEL,
                             "border": 1, "align": "center", "valign": "vcenter", "num_format": "0"})
    f_status = wb.add_format({"font_size": 11, "bold": True, "font_color": NAVY, "bg_color": "#EAF2FA",
                              "border": 1, "border_color": "#B9CEE0", "align": "left", "valign": "vcenter",
                              "text_wrap": True, "indent": 1})
    f_gridhdr = wb.add_format({"bold": True, "font_color": "white", "bg_color": STEEL, "align": "center",
                               "valign": "vcenter", "border": 1, "border_color": "white"})
    f_metric = wb.add_format({"font_size": 11, "font_color": "#22303C", "align": "left", "valign": "vcenter",
                              "border": 1, "border_color": "#DCE4EC", "indent": 1, "bg_color": LIGHT})
    f_pct = wb.add_format({"num_format": "0.00%", "align": "center", "valign": "vcenter", "border": 1,
                           "border_color": "#DCE4EC"})
    f_pct_blend = wb.add_format({"num_format": "0.00%", "align": "center", "valign": "vcenter", "border": 1,
                                 "border_color": "#DCE4EC", "bold": True, "font_color": NAVY, "bg_color": "#FFF6DD"})
    f_cur = wb.add_format({"num_format": '\u20b9 #,##0', "align": "center", "valign": "vcenter", "border": 1,
                           "border_color": "#DCE4EC", "bold": True})
    f_cur_blend = wb.add_format({"num_format": '\u20b9 #,##0', "align": "center", "valign": "vcenter", "border": 1,
                                 "border_color": "#DCE4EC", "bold": True, "font_color": NAVY, "bg_color": "#FFF6DD"})
    f_small = wb.add_format({"font_size": 9, "font_color": GREY, "align": "left", "valign": "vcenter"})
    f_meta_val = wb.add_format({"font_size": 10, "bold": True, "font_color": NAVY, "align": "left", "valign": "vcenter"})
    f_meta_dt = wb.add_format({"font_size": 10, "bold": True, "font_color": NAVY, "align": "left",
                               "valign": "vcenter", "num_format": "dd-mmm-yyyy hh:mm"})

    f_dbhdr = wb.add_format({"bold": True, "font_color": "white", "bg_color": SLATE, "align": "center",
                             "valign": "vcenter", "border": 1, "text_wrap": True})
    f_dbdate = wb.add_format({"num_format": "dd-mmm-yyyy", "align": "center"})
    f_dbnum = wb.add_format({"num_format": "#,##0.00", "align": "center"})

    # =====================================================================
    # Dashboard
    # =====================================================================
    dash = wb.add_worksheet("Dashboard")
    dash.set_vba_name("Sheet1")
    dash.hide_gridlines(2)
    dash.set_column("A:A", 2.5)
    dash.set_column("B:B", 32)
    dash.set_column("C:C", 16)
    dash.set_column("D:D", 16)
    dash.set_column("E:G", 15)
    dash.set_column("H:H", 3)
    dash.set_row(0, 40)
    dash.set_row(1, 20)

    dash.merge_range("A1:H1", "FACTOR INVESTMENT SIMULATOR  \u2014  ROLLING PRO", f_title)
    dash.merge_range("A2:H2", "NSE factor-cycle analysis  \u00b7  Momentum vs Value vs Benchmark vs custom blend  \u00b7  SIP (XIRR) & Lumpsum (CAGR)", f_sub)

    dash.merge_range("B4:D4", "INPUT PARAMETERS", f_section)

    inputs = [
        (5, "Investment Mode", "SIP", f_input),
        (6, "Investment Amount (\u20b9)", 25000, f_input_cur),
        (7, "Rolling Horizon (Years)", 5, f_input),
        (8, "Sample Start Date", date(2012, 1, 1), f_input_date),
        (9, "Sample End Date", date(2024, 12, 31), f_input_date),
        (10, "Allocation \u2014 Momentum (%)", 60, f_input),
        (11, "Allocation \u2014 Value (%)", 40, f_input),
    ]
    for r, label, val, fmt in inputs:
        dash.write(r - 1, 1, label, f_label)
        if isinstance(val, date):
            dash.write_datetime(r - 1, 3, val, fmt)
        else:
            dash.write(r - 1, 3, val, fmt)

    dash.write(11, 1, "Allocation Total (%)  \u2014  must equal 100", f_label)
    dash.write_formula(11, 3, "=D10+D11", f_total)

    # Dropdown validations
    dash.data_validation("D5", {"validate": "list", "source": ["SIP", "Lumpsum"],
                                "input_title": "Investment Mode",
                                "input_message": "Choose SIP or Lumpsum"})
    dash.data_validation("D7", {"validate": "list", "source": [3, 5, 10],
                                "input_title": "Rolling Horizon",
                                "input_message": "Choose 3, 5 or 10 years"})
    dash.data_validation("D10", {"validate": "integer", "criteria": "between", "minimum": 0, "maximum": 100})
    dash.data_validation("D11", {"validate": "integer", "criteria": "between", "minimum": 0, "maximum": 100})

    # Status
    dash.write(13, 1, "STATUS", f_section)
    dash.merge_range("D14:H14", "Ready. Enter inputs, then click Run Simulation. Use Refresh Data to import NSE files.", f_status)
    dash.set_row(13, 30)

    # Results grid
    dash.merge_range("B16:G16", "RESULTS  \u2014  ROLLING RETURN COMPARISON", f_section)
    dash.merge_range("B17:C17", "Metric", f_gridhdr)
    for c, name in zip((3, 4, 5, 6), ("Momentum", "Value", "Benchmark", "Blended")):
        dash.write(16, c, name, f_gridhdr)

    metrics = [
        (18, "Average Rolling Return"),
        (19, "Worst-Case Rolling Floor"),
        (20, "Best-Case Rolling Peak"),
        (21, "Probability of Return > 12%"),
        (22, "Probability of Return > 15%"),
        (23, "Final Projected Maturity Corpus (\u20b9)"),
    ]
    for r, label in metrics:
        dash.merge_range(r - 1, 1, r - 1, 2, label, f_metric)
        for c in (3, 4, 5, 6):
            if r == 23:
                dash.write_blank(r - 1, c, None, f_cur_blend if c == 6 else f_cur)
            else:
                dash.write_blank(r - 1, c, None, f_pct_blend if c == 6 else f_pct)

    # Run metadata (VBA writes B25 count and B26 timestamp)
    dash.write_blank("B25", None, f_meta_val)
    dash.write("C25", "rolling windows evaluated", f_small)
    dash.write_blank("B26", None, f_meta_dt)
    dash.write("C26", "last run (date / time)", f_small)

    # Buttons wired to macros
    dash.insert_button("F5", {"macro": "RefreshData", "caption": "Refresh Data",
                              "width": 132, "height": 30})
    dash.insert_button("F7", {"macro": "RunSimulation", "caption": "Run Simulation",
                              "width": 132, "height": 30})
    dash.insert_button("F9", {"macro": "ResetInputs", "caption": "Reset Inputs",
                              "width": 132, "height": 30})

    # =====================================================================
    # Database_Daily
    # =====================================================================
    db = wb.add_worksheet("Database_Daily")
    db.set_column("A:A", 14)
    db.set_column("B:D", 20)
    db.freeze_panes(1, 0)
    headers = ["Date",
               "Momentum_TRI (Nifty 200 Momentum 30)",
               "Value_TRI (Nifty 500 Value 50)",
               "Benchmark_TRI (Nifty 500)"]
    for c, h in enumerate(headers):
        db.write(0, c, h, f_dbhdr)
    db.set_row(0, 30)
    for i, d in enumerate(dates):
        row = i + 1
        db.write_datetime(row, 0, d, f_dbdate)
        db.write_number(row, 1, round(series["mom"][i], 2), f_dbnum)
        db.write_number(row, 2, round(series["val"][i], 2), f_dbnum)
        db.write_number(row, 3, round(series["ben"][i], 2), f_dbnum)

    # =====================================================================
    # Config
    # =====================================================================
    cfg = wb.add_worksheet("Config")
    cfg.hide_gridlines(2)
    cfg.set_column("A:A", 2.5)
    cfg.set_column("B:B", 40)
    cfg.set_column("C:C", 2)
    cfg.set_column("D:D", 60)
    cfg.set_row(0, 34)
    cfg.merge_range("A1:D1", "CONFIGURATION  \u2014  DATA IMPORT & ASSUMPTIONS", f_title)

    cfg.merge_range("B3:D3", "IMPORT FILE PATHS (official NSE / NSE Indices CSV exports)", f_section)
    cfg_rows = [
        (5, "Nifty 200 Momentum 30 TRI  \u2014  CSV path"),
        (6, "Nifty 500 Value 50 TRI  \u2014  CSV path"),
        (7, "Nifty 500 Parent Benchmark TRI  \u2014  CSV path"),
    ]
    for r, label in cfg_rows:
        cfg.write(r - 1, 1, label, f_label)
        cfg.write_blank(r - 1, 3, None, f_input)
    cfg.write(8, 1, "Example (Windows):  C:\\NSE\\momentum30_tri.csv", f_small)

    cfg.merge_range("B10:D10", "HOW REFRESH WORKS / ASSUMPTIONS", f_section)
    notes = [
        "\u2022 Leave a path blank to keep the existing column unchanged for that index.",
        "\u2022 The Date column and the TRI/value column are auto-detected from each file's header row.",
        "\u2022 Accepted date formats include 01-Jan-2024, 01 Jan 2024, 01/01/2024 and 2024-01-01.",
        "\u2022 All three histories are merged by date, sorted chronologically and forward-filled for gaps.",
        "\u2022 If no files are found, seeded sample history is retained (nothing is deleted).",
        "\u2022 Rolling windows are generated at a monthly cadence within the selected sample period.",
        "\u2022 Lumpsum uses exact CAGR; SIP uses exact XIRR computed natively in VBA (Newton-Raphson).",
        "\u2022 Blended portfolio applies a fixed Momentum/Value split at every contribution \u2014 no rebalancing.",
    ]
    for i, n in enumerate(notes):
        cfg.write(11 + i, 1, n, f_label)

    # =====================================================================
    # Calc_Cache (hidden, audit/debug)
    # =====================================================================
    cache = wb.add_worksheet("Calc_Cache")
    cache.set_column("A:B", 14)
    cache.set_column("C:F", 12)
    ch = ["Entry Date", "Maturity Date", "Momentum", "Value", "Benchmark", "Blended"]
    for c, h in enumerate(ch):
        cache.write(0, c, h, f_dbhdr)
    cache.set_column("A:B", 14, wb.add_format({"num_format": "dd-mmm-yyyy"}))
    cache.set_column("C:F", 12, wb.add_format({"num_format": "0.00%"}))
    cache.hide()

    # ---- Named ranges ----
    names = {
        "in_Mode": "'Dashboard'!$D$5",
        "in_Amount": "'Dashboard'!$D$6",
        "in_Horizon": "'Dashboard'!$D$7",
        "in_StartDate": "'Dashboard'!$D$8",
        "in_EndDate": "'Dashboard'!$D$9",
        "in_AllocMom": "'Dashboard'!$D$10",
        "in_AllocVal": "'Dashboard'!$D$11",
        "out_Status": "'Dashboard'!$D$14",
        "cfg_MomFile": "'Config'!$D$5",
        "cfg_ValFile": "'Config'!$D$6",
        "cfg_BenFile": "'Config'!$D$7",
    }
    for nm, ref in names.items():
        wb.define_name(nm, "=" + ref)

    wb.add_vba_project(vba_bin)
    dash.activate()
    wb.close()
    print("Workbook written:", OUT, "| rows:", len(dates))


if __name__ == "__main__":
    b = build_vba_project()
    build_workbook(b)
    # verify
    with zipfile.ZipFile(OUT) as z:
        assert "xl/vbaProject.bin" in z.namelist(), "vbaProject.bin missing!"
    print("Final VBA modules:", ExcelFile(OUT).module_names())
    print("SIZE (bytes):", os.path.getsize(OUT))
