#!/usr/bin/env python3
"""Build an Excel workbook of all indices (Price + TRI) from the latest snapshot.
Two sheets: 'Price (Unadjusted)' and 'TRI (Total Return)'. Run after refresh_data.py."""
import json, os
from datetime import date
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
SNAP = os.path.join(HERE, "data_snapshots", "latest.json")
OUT = os.path.join(ROOT, "NSE_Index_Data_Price_TRI.xlsx")

data = json.load(open(SNAP, encoding="utf-8"))
dates, order, labels = data["dates"], data["order"], data["labels"]

wb = Workbook()
hdr_fill = PatternFill("solid", fgColor="0F2A43")
hdr_font = Font(bold=True, color="FFFFFF", size=11)
date_font = Font(bold=True)


def build_sheet(ws, series):
    ws.append(["Date"] + [labels[k] for k in order])
    for ci in range(1, len(order) + 2):
        c = ws.cell(row=1, column=ci)
        c.fill = hdr_fill; c.font = hdr_font; c.alignment = Alignment(horizontal="center")
    for i, d in enumerate(dates):
        ws.append([d] + [series[k][i] for k in order])
    ws.freeze_panes = "B2"
    ws.column_dimensions["A"].width = 13
    for j in range(len(order)):
        ws.column_dimensions[get_column_letter(j + 2)].width = max(12, len(labels[order[j]]) + 2)
    for r in range(2, len(dates) + 2):
        ws.cell(row=r, column=1).font = date_font
        for j in range(len(order)):
            ws.cell(row=r, column=j + 2).number_format = "#,##0.00"


ws1 = wb.active; ws1.title = "Price (Unadjusted)"
build_sheet(ws1, data["price"])
ws2 = wb.create_sheet("TRI (Total Return)")
build_sheet(ws2, data["tri"])

wb.save(OUT)
print("wrote", OUT)
print("indices:", ", ".join(labels[k] for k in order))
print("rows:", len(dates), "range:", dates[0], "->", dates[-1])
