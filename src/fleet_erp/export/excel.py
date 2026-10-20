"""
Export a Projection to a structured Excel workbook that mirrors the original
output: SV_PLAN_v1 (calendar), SV_PLAN_v2 (event × aircraft), one sheet per aircraft.
"""

from __future__ import annotations

import io
import re

import pandas as pd

_ILLEGAL = re.compile(r"[:\\/?*\[\]]")


def _safe(name: str, used: set) -> str:
    n = _ILLEGAL.sub("", str(name))[:31] or "sheet"
    base, i = n, 1
    while n in used:
        suf = f"_{i}"
        n = base[:31 - len(suf)] + suf
        i += 1
    used.add(n)
    return n


def build_workbook(proj) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter", datetime_format="mmm-yy") as xw:
        wb = xw.book
        hdr = wb.add_format({"bold": True, "font_color": "white", "bg_color": "#1F3864",
                             "border": 1, "align": "center", "valign": "vcenter"})
        red = wb.add_format({"bg_color": "#FFC7CE", "bold": True})
        used: set = set()

        def write(df, sheet, highlight_zero=False):
            name = _safe(sheet, used)
            df.to_excel(xw, sheet_name=name, index=False)
            ws = xw.sheets[name]
            for ci, col in enumerate(df.columns):
                ws.write(0, ci, col, hdr)
                width = 12 if ci else 16
                ws.set_column(ci, ci, width)
            ws.freeze_panes(1, 1)
            # highlight shop-visit months (SLSV == 0) on per-aircraft sheets
            if highlight_zero:
                for ci, col in enumerate(df.columns):
                    if "SLSV" not in str(col):
                        continue
                    for ri in range(len(df)):
                        v = df.iloc[ri, ci]
                        if pd.notna(v) and float(v) == 0.0:
                            ws.write(ri + 1, ci, 0, red)

        if proj.sv_plan_v1 is not None and not proj.sv_plan_v1.empty:
            write(proj.sv_plan_v1, "SV_PLAN_v1")
        if proj.sv_plan_v2 is not None and not proj.sv_plan_v2.empty:
            write(proj.sv_plan_v2, "SV_PLAN_v2")
        for tail, frame in (proj.aircraft_frames or {}).items():
            write(frame, tail, highlight_zero=True)

    buf.seek(0)
    return buf.getvalue()
