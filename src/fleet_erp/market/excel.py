"""Excel export of the Market Intelligence report — the management "Lessors Update" workbook."""
from __future__ import annotations

import io

import xlsxwriter

_NAVY, _BLUE, _TEAL, _AMBER, _GREY = "#12233F", "#1F5FBF", "#0F766E", "#B45309", "#EEF2F7"


def build(rep: dict, fit: dict | None, meta: dict) -> bytes:
    buf = io.BytesIO()
    wb = xlsxwriter.Workbook(buf, {"in_memory": True})
    f = {
        "title": wb.add_format({"bold": True, "font_size": 14, "font_color": _NAVY}),
        "sub": wb.add_format({"italic": True, "font_color": "#64748B"}),
        "txt": wb.add_format({"border": 1, "border_color": "#D5DCE6"}),
        "num": wb.add_format({"border": 1, "border_color": "#D5DCE6", "num_format": "#,##0;-#,##0;"}),
        "pct": wb.add_format({"border": 1, "border_color": "#D5DCE6", "num_format": '0.0"%"'}),
        "ttl": wb.add_format({"border": 1, "border_color": "#D5DCE6", "num_format": "#,##0;-#,##0;",
                              "bold": True, "bg_color": _GREY}),
        "tot_txt": wb.add_format({"bold": True, "bg_color": _NAVY, "font_color": "white", "border": 1}),
        "tot_num": wb.add_format({"bold": True, "bg_color": _NAVY, "font_color": "white", "border": 1,
                                  "num_format": "#,##0;-#,##0;"}),
        "delta": wb.add_format({"border": 1, "border_color": "#D5DCE6", "num_format": "+#,##0;-#,##0;"}),
    }

    def head(color):
        return wb.add_format({"bold": True, "bg_color": color, "font_color": "white", "border": 1,
                              "align": "center", "valign": "vcenter", "text_wrap": True})

    h_navy, h_blue, h_teal, h_amber = head(_NAVY), head(_BLUE), head(_TEAL), head(_AMBER)

    fams, ofams, rtypes = rep["families"], rep["orderFamilies"], rep["relTypes"]
    rel = rep.get("relationship") or {}
    airline = rel.get("airline") or "airline"

    # ── Summary 2: the lessor league ─────────────────────────────────────────
    ws = wb.add_worksheet("Lessor table")
    ws.write(0, 0, "Lessors Update — market overview", f["title"])
    ws.write(1, 0, f"Source: {meta.get('file', 'Asset report')} · generated {meta.get('at', '')}", f["sub"])
    r0, c = 3, 0
    ws.merge_range(r0, 0, r0 + 1, 0, "Manager (not necessarily owner)", h_navy)
    ws.merge_range(r0, 1, r0 + 1, 1, "Funding / main shareholders", h_navy)
    c = 2
    blocks = []
    if rtypes:
        blocks.append((f"Current relationship with {airline}", rtypes + ["TTL"], h_amber, "rel"))
    blocks.append(("Current fleet", fams + ["TTL"], h_blue, "active"))
    if rep["hasPrev"]:
        blocks.append(("vs. previous", ["Previous", "Δ"], h_navy, "prev"))
    blocks.append(("Orders", ofams + ["TTL"], h_teal, "orders"))
    layout = []
    for title, cols, fmt, key in blocks:
        if len(cols) > 1:
            ws.merge_range(r0, c, r0, c + len(cols) - 1, title, fmt)
        else:
            ws.write(r0, c, title, fmt)
        for j, name in enumerate(cols):
            ws.write(r0 + 1, c + j, name, fmt)
        layout.append((key, cols, c))
        c += len(cols)

    def write_row(r, row, total=False):
        tf, nf = (f["tot_txt"], f["tot_num"]) if total else (f["txt"], f["num"])
        ws.write(r, 0, row["name"], tf)
        ws.write(r, 1, row.get("funding", ""), tf)
        for key, cols, c0 in layout:
            if key == "prev":
                ws.write(r, c0, row.get("prevTotal") or 0, nf)
                ws.write(r, c0 + 1, row.get("delta") or 0, nf if total else f["delta"])
                continue
            src = row.get(key, {})
            for j, name in enumerate(cols[:-1]):
                ws.write(r, c0 + j, src.get(name, 0), nf)
            ttl = {"rel": "relTotal", "active": "activeTotal", "orders": "ordersTotal"}[key]
            ws.write(r, c0 + len(cols) - 1, row.get(ttl, 0), nf if total else f["ttl"])

    r = r0 + 2
    for row in rep["lessors"]:
        write_row(r, row)
        r += 1
    write_row(r, rep["lessorTotal"], total=True)
    write_row(r + 1, rep["other"])
    write_row(r + 2, rep["total"], total=True)
    r += 4
    if rel:
        ws.write(r, 0, f"{airline} — fleet outside the listed lessors", f["title"])
        r += 1
        extra = [("OWNED", rel.get("owned", {}))] + [(u["lessor"], u["byType"]) for u in rel.get("unmatched", [])]
        for name, by in extra:
            if not by:
                continue
            ws.write(r, 0, name, f["txt"])
            ws.write(r, 1, "", f["txt"])
            for j, t in enumerate(rtypes):
                ws.write(r, 2 + j, by.get(t, 0), f["num"])
            ws.write(r, 2 + len(rtypes), sum(by.values()), f["ttl"])
            r += 1
    ws.set_column(0, 0, 38)
    ws.set_column(1, 1, 44)
    ws.set_column(2, c, 9)
    ws.set_row(r0, 30)
    ws.freeze_panes(r0 + 2, 2)

    def table(name, title, columns, rows, widths=None):
        s = wb.add_worksheet(name)
        s.write(0, 0, title, f["title"])
        for j, col in enumerate(columns):
            s.write(2, j, col, h_navy)
        for i, row in enumerate(rows):
            for j, v in enumerate(row):
                s.write(3 + i, j, v, f["num"] if isinstance(v, (int, float)) else f["txt"])
        for j, w in enumerate(widths or []):
            s.set_column(j, j, w)
        s.freeze_panes(3, 1)
        return s

    by_orders = sorted(rep["lessors"], key=lambda x: -x["ordersTotal"])
    table("Orderbook", "Orderbook ranking — listed lessors",
          ["#", "Lessor", "Funding"] + ofams + ["Orders TTL", "Active fleet", "Backlog % of fleet", "Share of all orders %"],
          [[i, x["name"], x["funding"]] + [x["orders"].get(fm, 0) for fm in ofams]
           + [x["ordersTotal"], x["activeTotal"], x["backlog"] or 0, x["orderShare"]]
           for i, x in enumerate([x for x in by_orders if x["ordersTotal"]], 1)],
          [5, 38, 44] + [9] * len(ofams) + [11, 11, 13, 13])

    table("Families", "Aircraft families — market vs. listed lessors",
          ["Family", "Active", "On order", "Lessor-managed", "Lessor share %", "Lessor orders",
           "Lessor order share %", "Avg age (yrs)", "Stored / parked", "Largest lessor", "Aircraft",
           "Largest orderbook", "Orders"],
          [[x["family"], x["active"], x["onOrder"], x["lessorActive"], x["lessorShare"] or 0, x["lessorOrders"],
            x["lessorOrderShare"] or 0, x["avgAge"] or 0, x["idle"], x["topLessor"], x["topLessorCount"],
            x["topOrderLessor"], x["topOrderCount"]] for x in rep["familyTable"]],
          [12, 10, 10, 14, 13, 13, 16, 12, 14, 34, 9, 34, 9])

    table("Capital origin", "Where the lessors' capital comes from",
          ["Origin", "Lessors", "Active fleet", "Share of lessor fleet %", "Orders", "Largest"],
          [[x["country"], x["lessors"], x["active"], x["share"], x["orders"], ", ".join(x["names"])]
           for x in rep["origins"]], [14, 9, 12, 18, 9, 90])

    if fit:
        table("Fleet fit", f"Fleet fit — {fit.get('airline') or 'airline'}",
              ["Family", "Airline aircraft", "Share of fleet %", "Successor", "Market active", "Market on order",
               "Lessor-managed", "Lessor orders", "Stored / parked", "Market avg age"],
              [[x["family"], x["airline"], x["airlineShare"], x["successor"], x["marketActive"], x["marketOrders"],
                x["lessorActive"], x["lessorOrders"], x["idle"], x["marketAvgAge"] or 0] for x in fit["families"]],
              [12, 14, 14, 11, 12, 14, 14, 12, 14, 14])
        tg = fit["targets"]
        table("Sourcing shortlist", "Who could supply the target families: " + ", ".join(tg),
              ["Lessor", "Funding", "Existing relationship", "Aircraft with airline"]
              + [f"{t} active" for t in tg] + [f"{t} orders" for t in tg] + ["Stored / parked", "Score"],
              [[x["lessor"], x["funding"], "Yes" if x["existing"] else "No", x["relTotal"]]
               + [x["active"].get(t, 0) for t in tg] + [x["orders"].get(t, 0) for t in tg]
               + [x["idle"], x["score"]] for x in fit["shortlist"]],
              [38, 44, 12, 12] + [12] * (2 * len(tg)) + [12, 9])

    ql = rep["quality"]
    table("Settings gaps", "Third-party managers not in Settings (counted under Other)",
          ["Manager", "Active", "On order", "Operators served"],
          [[x["name"], x["active"], x["orders"], x["operators"]] for x in ql["unmappedManagers"]],
          [44, 10, 10, 14])
    wb.close()
    return buf.getvalue()
