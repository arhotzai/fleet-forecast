"""
Fleet Forecast — REST API (Phase 1, no database).

A thin FastAPI layer over the proven Python engine in ``src/fleet_erp``.
Endpoints and JSON shapes match the Lovable React UI's contract
(src/types.ts / src/lib/api.ts): a single "current" dataset, no session id.

IMPORTANT: nothing about the fleet is hardcoded. Unit kinds, types, events,
drivers and limits are all read from the uploaded file — the file dictates
everything, exactly like the Streamlit app.
"""
from __future__ import annotations

import io
import os
import sys

import numpy as np
import pandas as pd
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(_ROOT, "src"))

from fleet_erp.ingestion.loader import load          # noqa: E402
from fleet_erp.technical import engine               # noqa: E402
from fleet_erp.export.excel import build_workbook     # noqa: E402

app = FastAPI(title="Fleet Forecast API", version="1.0")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_credentials=False,
                   allow_methods=["*"], allow_headers=["*"])

# single in-memory dataset (Phase 1, no DB). None until a workbook is uploaded.
STATE: dict = {"ds": None, "proj": None, "airline": ""}
_SAMPLE = os.path.join(_ROOT, "assets", "SAMPLE_fleet_database.xlsx")
_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# load backend secrets (ANTHROPIC_API_KEY, …) from .env.local if present — gitignored
_ENV = os.path.join(_ROOT, ".env.local")
if os.path.exists(_ENV):
    for _line in open(_ENV):
        _line = _line.strip()
        if _line and not _line.startswith("#") and "=" in _line:
            _k, _v = _line.split("=", 1)
            os.environ.setdefault(_k.strip(), _v.strip().strip('"').strip("'"))


# ───────────────────────── json helpers ─────────────────────────
def _num(v, default=0):
    try:
        f = float(v)
    except (TypeError, ValueError):
        return default
    if pd.isna(f):
        return default
    return int(f) if f == int(f) else round(f, 2)


def _numN(v):   # nullable number
    r = _num(v, None)
    return r


def _str(v, default=""):
    if v is None:
        return default
    try:
        if pd.isna(v):
            return default
    except (TypeError, ValueError):
        pass
    return str(v).strip() or default


def _date(v, default=""):
    try:
        if pd.isna(v):
            return default
    except (TypeError, ValueError):
        pass
    try:
        return pd.to_datetime(v).strftime("%Y-%m-%d")
    except (TypeError, ValueError):
        return _str(v, default)


# ───────────────────── display mappers (data-driven) ─────────────────────
def _kind(k: str) -> str:
    u = str(k).upper()
    return {"AIRFRAME": "Airframe", "ENGINE": "Engine", "APU": "APU",
            "PROPELLER": "Propeller"}.get(u, u.replace("_", " ").title())


def _category(k: str) -> str:
    u = str(k).upper()
    return {"AIRFRAME": "Airframe check", "ENGINE": "Engine", "APU": "APU",
            "PROPELLER": "Propeller"}.get(u, u.replace("_", " ").title())


def _driver(d: str) -> str:
    u = str(d).upper()
    if u.startswith("CAL") or u == "CLND":
        return "Calendar"
    return "FC" if u == "FC" else "FH"


def _reason(trigger: str) -> str:
    return {"LLP_REPLACEMENT": "llp", "HARD_LINK": "hard",
            "SOFT_BUNDLE": "soft", "LIMIT": "limit"}.get(str(trigger), "limit")


def _big_cat(name: str) -> str:
    f = str(name).upper()
    if "NEO" in f:
        return "NEO"
    if "CEO" in f:
        return "CEO"
    if "ATR" in f or "DASH" in f or "DH" in f or "Q400" in f:
        return "ATR"
    return "Other"


_ENG_CAT = {"PW1100": "NEO", "PW1500": "NEO", "LEAP": "NEO",
            "V2500": "CEO", "CFM56": "CEO", "PW127": "ATR"}
_POSL = {"LHENG": "LH", "RHENG": "RH", "APU": "APU", "LHPROP": "LH", "RHPROP": "RH",
         "AIRFRAME": "—", "SPARE": "Spare"}


def _ds():
    if STATE["ds"] is None:
        return None
    return STATE["ds"]


def _is_spare(tail) -> bool:
    return tail is None or str(tail).strip() in ("", "nan", "None")


# ───────────────────── derived lookups (per request) ─────────────────────
def _tail_family(ds) -> dict:
    fleet = ds.get("Fleet")
    out = {}
    if fleet is not None and not fleet.empty:
        for _, r in fleet.iterrows():
            out[str(r["Tail"])] = _big_cat(r.get("Aircraft_Family"))
    return out


def _util_by_tail(ds) -> dict:
    u = ds.get("Utilization")
    out = {}
    if u is not None and not u.empty:
        for _, r in u.iterrows():
            out[str(r["Unit_ID"])] = (_num(r.get("Annual_FH")), _num(r.get("Annual_FC")))
    return out


def _snapshot_date(ds):
    cfg = ds.get("Config")
    if cfg is not None and not cfg.empty and "Setting" in cfg.columns:
        m = cfg[cfg["Setting"].astype(str) == "As_Of_Date"]
        if not m.empty:
            return _date(m.iloc[0]["Value"])
    return ""


def _active_today(fl):
    if fl is None or fl.empty:
        return fl
    d = fl.copy()
    for c in ("Delivery_Date", "Redelivery_Date", "Extension_Date"):
        if c in d.columns:
            d[c] = pd.to_datetime(d[c], errors="coerce")
    today = pd.Timestamp.today().normalize()
    end = d["Redelivery_Date"] if "Redelivery_Date" in d.columns else None
    if "Extension_Date" in d.columns:
        end = d["Extension_Date"].fillna(end) if end is not None else d["Extension_Date"]
    m = pd.Series(True, index=d.index)
    if "Delivery_Date" in d.columns:
        m &= d["Delivery_Date"].isna() | (d["Delivery_Date"] <= today)
    if end is not None:
        m &= end.isna() | (end >= today)
    return d[m]


# ───────────────────────── builders ─────────────────────────
def _aircraft_list(ds) -> list[dict]:
    fleet = ds.get("Fleet")
    if fleet is None or fleet.empty:
        return []
    util = _util_by_tail(ds)
    out = []
    for _, r in fleet.iterrows():
        tail = str(r["Tail"])
        fh, fc = util.get(tail, (0, 0))
        out.append({
            "tail": tail, "msn": _num(r.get("MSN")),
            "family": _big_cat(r.get("Aircraft_Family")),
            "type": _str(r.get("Aircraft_Family")),
            "owner": _str(r.get("Operator")), "lessor": _str(r.get("Owner")),
            "leasedOwned": _str(r.get("Leased_Owned")),
            "deliveryDate": _date(r.get("Delivery_Date")),
            "fhYear": fh, "fcYear": fc,
        })
    return out


def _status_rows(ds) -> list[dict]:
    us = ds.get("Unit_Status")
    comp = ds.get("Components")
    evdef = ds.get("Events")
    fleet = ds.get("Fleet")
    if us is None or us.empty:
        return []
    info = {}
    if comp is not None and not comp.empty:
        for _, r in comp.iterrows():
            info[str(r["Component_ID"])] = (str(r["Unit_Kind"]), str(r["Unit_Type"]),
                                            str(r.get("Position") or ""),
                                            str(r.get("Installed_On_Tail") or ""))
    if fleet is not None and not fleet.empty:
        for _, r in fleet.iterrows():
            info[str(r["Tail"])] = ("AIRFRAME", str(r.get("Aircraft_Family") or ""),
                                    "AIRFRAME", str(r["Tail"]))
    lim = {}
    if evdef is not None and not evdef.empty:
        for _, r in evdef.iterrows():
            lim[(str(r["Unit_Type"]), str(r["Event_ID"]))] = (r.get("Hard_Limit"),
                                                              str(r.get("Driver") or ""))
    rows = []
    for _, r in us.iterrows():
        uid = str(r["Unit_ID"])
        inf = info.get(uid)
        if not inf:
            continue
        kind, utype, pos, tail = inf
        evid = str(r["Event_ID"])
        hl, drv = lim.get((utype, evid), (None, str(r.get("Driver") or "")))
        drv = drv or str(r.get("Driver") or "")
        tsn = pd.to_numeric(r.get("TSN"), errors="coerce")
        csn = pd.to_numeric(r.get("CSN"), errors="coerce")
        tsn_l = pd.to_numeric(r.get("TSN_At_LSV"), errors="coerce")
        csn_l = pd.to_numeric(r.get("CSN_At_LSV"), errors="coerce")
        slsv_fh = pd.to_numeric(r.get("SLSV_FH"), errors="coerce")
        slsv_fc = pd.to_numeric(r.get("SLSV_FC"), errors="coerce")
        cyc = (csn - csn_l) if (pd.notna(csn) and pd.notna(csn_l) and csn_l) else slsv_fc
        if drv.startswith("CAL") or drv == "CLND":
            consumed = pd.to_numeric(r.get("SLSV_Months"), errors="coerce")
        elif drv == "FC":
            consumed = cyc
        else:
            consumed = (tsn - tsn_l) if (pd.notna(tsn) and pd.notna(tsn_l) and tsn_l) else slsv_fh
        hln = pd.to_numeric(hl, errors="coerce")
        remaining = (hln - consumed) if (pd.notna(hln) and pd.notna(consumed)) else None
        pct = round(consumed / hln * 100, 1) if (pd.notna(hln) and hln and pd.notna(consumed)) else None
        rows.append({
            "id": f"{uid}|{evid}", "tail": _str(tail), "unit": uid,
            "kind": _kind(kind), "type": utype, "pos": _POSL.get(pos, pos),
            "event": evid, "driver": _driver(drv),
            "sinceLast": _num(consumed), "cyclesSinceSv": _num(cyc),
            "limit": _num(hln), "remaining": _num(remaining), "pctLimit": _num(pct),
            "tsnLsv": _num(tsn_l), "csnLsv": _num(csn_l),
            "tsnNow": _num(tsn), "csnNow": _num(csn), "lastSv": _date(r.get("Last_SV_Date")),
        })
    rows.sort(key=lambda x: -(x["pctLimit"] or 0))
    return rows


def _power_units(ds) -> list[dict]:
    comp = ds.get("Components")
    us = ds.get("Unit_Status")
    llp = ds.get("LLP_Registry")
    proj = STATE["proj"]
    if comp is None or comp.empty:
        return []
    pw = comp[comp["Unit_Kind"].isin(["ENGINE", "APU"])]
    posmap = {"LHENG": "LH", "RHENG": "RH", "APU": "APU", "SPARE": "Spare"}
    agg = {}
    if us is not None and not us.empty:
        acols = {"tsn": ("TSN", "first"), "csn": ("CSN", "first"),
                 "lastSvDate": ("Last_SV_Date", "first")}
        if "TSN_At_LSV" in us.columns:
            acols["tsnLsv"] = ("TSN_At_LSV", "first")
        if "WS_Type" in us.columns:
            acols["workscope"] = ("WS_Type", "first")
        agg = us.groupby(us["Unit_ID"].astype(str)).agg(**acols).to_dict("index")
    lc = llp.groupby(llp["Unit_ID"].astype(str)).size().to_dict() if (llp is not None and not llp.empty) else {}
    tfam = _tail_family(ds)
    # projected shop visits per unit
    sv = {}
    if proj is not None and proj.events is not None and not proj.events.empty:
        for _, e in proj.events.iterrows():
            sv.setdefault(str(e["Unit_ID"]), []).append({
                "date": _date(e.get("Date")), "workscope": _str(e.get("Event")),
                "tsn": 0, "csn": 0, "vendor": ""})
    out = []
    for _, r in pw.iterrows():
        cid = str(r["Component_ID"])
        pos = str(r.get("Position") or "")
        tail = None if _is_spare(r.get("Installed_On_Tail")) else str(r.get("Installed_On_Tail"))
        a = agg.get(cid, {})
        utype = _str(r.get("Unit_Type"))
        fam = tfam.get(tail) if tail else _ENG_CAT.get(utype, "Other")
        out.append({
            "serial": cid, "type": utype, "kind": _kind(r["Unit_Kind"]),
            "family": fam or "Other", "position": posmap.get(pos, pos),
            "state": "Spare" if not tail else "Installed", "onTail": tail,
            "tsn": _num(a.get("tsn")), "tsnLsv": _num(a.get("tsnLsv")),
            "csn": _num(a.get("csn")), "lastSvDate": _date(a.get("lastSvDate")),
            "workscope": _str(a.get("workscope")), "llpCount": int(lc.get(cid, 0)),
            "shopVisits": sv.get(cid, []),
        })
    return out


def _events(ds) -> list[dict]:
    proj = STATE["proj"]
    if proj is None or proj.events is None or proj.events.empty:
        return []
    tfam = _tail_family(ds)
    out = []
    for i, e in proj.events.reset_index(drop=True).iterrows():
        raw_tail = _str(e.get("Tail"))
        unit = _str(e.get("Unit_ID"))
        is_spare = raw_tail in ("", "SPARE")
        tail = unit if is_spare else raw_tail      # spares shown by their own serial
        fam = "Spares" if is_spare else (tfam.get(raw_tail) or _ENG_CAT.get(_str(e.get("Unit_Type"))) or "Other")
        out.append({
            "id": str(i), "tail": tail, "unit": unit,
            "kind": _kind(e.get("Unit_Kind")), "family": fam,
            "event": _str(e.get("Event")) or _str(e.get("Event_ID")),
            "category": _category(e.get("Unit_Kind")),
            "dueDate": _date(e.get("Date")), "reason": _reason(e.get("Trigger")),
        })
    return out


def _llps(ds) -> list[dict]:
    reg = ds.get("LLP_Registry")
    comp = ds.get("Components")
    if reg is None or reg.empty:
        return []
    etype = {}
    if comp is not None and not comp.empty:
        etype = dict(zip(comp["Component_ID"].astype(str), comp["Unit_Type"].astype(str)))
    out = []
    for i, r in reg.reset_index(drop=True).iterrows():
        uid = _str(r.get("Unit_ID"))
        lim = _num(r.get("Limit_Cycles"))
        togo = _num(r.get("ToGo"))
        out.append({
            "id": str(i), "unit": uid, "engineType": etype.get(uid, ""),
            "part": _str(r.get("Description")), "pn": _str(r.get("LLP_PN")),
            "sn": _str(r.get("LLP_SN")), "limit": lim,
            "csn": (lim - togo) if (isinstance(lim, (int, float)) and isinstance(togo, (int, float))) else _num(r.get("SLSV")),
            "remaining": togo,
        })
    return out


def _utilisation(ds) -> list[dict]:
    fleet = ds.get("Fleet")
    util = ds.get("Utilization")
    if fleet is None or fleet.empty:
        return []
    um = {}
    if util is not None and not util.empty:
        for _, r in util.iterrows():
            um[str(r["Unit_ID"])] = (_num(r.get("Annual_FH")), _num(r.get("Annual_FC")),
                                     _num(r.get("FH_FC_Ratio")))
    rows = {}
    for _, r in fleet.iterrows():
        typ = _str(r.get("Aircraft_Family"))
        tail = str(r["Tail"])
        fh, fc, ratio = um.get(tail, (0, 0, 0))
        g = rows.setdefault(typ, {"family": _big_cat(typ), "type": typ, "count": 0,
                                  "_fh": 0.0, "_fc": 0.0, "_ratio": 0.0})
        g["count"] += 1
        g["_fh"] += fh or 0
        g["_fc"] += fc or 0
        g["_ratio"] += ratio or 0
    out = []
    for g in rows.values():
        n = max(g["count"], 1)
        out.append({"family": g["family"], "type": g["type"], "count": g["count"],
                    "fhYear": round(g["_fh"] / n), "fcYear": round(g["_fc"] / n),
                    "ratio": round(g["_ratio"] / n, 2)})
    return out


def _sheets(ds) -> list[dict]:
    out = []
    for name, df in ds.sheets.items():
        if df is None:
            continue
        cols = [str(c) for c in df.columns]
        rows = []
        for _, r in df.iterrows():
            row = []
            for c in df.columns:
                v = r[c]
                if isinstance(v, (int, float, np.integer, np.floating)):
                    row.append(_num(v, 0))
                else:
                    row.append(_str(v))
            rows.append(row)
        out.append({"name": name, "columns": cols, "rows": rows})
    return out


def _sheets_to_bytes(sheets) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as w:
        for sh in sheets:
            name = sh["name"][:31] if isinstance(sh, dict) else sh.name[:31]
            cols = sh["columns"] if isinstance(sh, dict) else sh.columns
            rows = sh["rows"] if isinstance(sh, dict) else sh.rows
            pd.DataFrame(rows, columns=cols).to_excel(w, sheet_name=name, index=False)
    return buf.getvalue()


# ───────────────────────── routes ─────────────────────────
@app.get("/")
def root():
    return {"service": "Fleet Forecast API", "status": "ok",
            "loaded": STATE["ds"] is not None, "docs": "/docs"}


@app.get("/snapshot")
def snapshot():
    ds = _ds()
    if ds is None:
        return {"date": "", "airline": "", "loaded": False}
    return {"date": _snapshot_date(ds), "airline": STATE["airline"], "loaded": True}


@app.get("/fleet/summary")
def summary():
    ds = _ds()
    if ds is None:
        return {"liveTotal": 0, "neo": 0, "ceo": 0, "atr": 0,
                "spareNeoEngines": 0, "spareCeoEngines": 0, "spareAtrEngines": 0, "spareApus": 0}
    active = _active_today(ds.get("Fleet"))
    cat = {"NEO": 0, "CEO": 0, "ATR": 0}
    if active is not None and not active.empty and "Aircraft_Family" in active.columns:
        for fam, cnt in active["Aircraft_Family"].value_counts().items():
            c = _big_cat(fam)
            if c in cat:
                cat[c] += int(cnt)
    total = int(active["Tail"].nunique()) if (active is not None and not active.empty) else 0
    sp = {"NEO": 0, "CEO": 0, "ATR": 0}
    n_apu = 0
    comp = ds.get("Components")
    if comp is not None and not comp.empty and "Installed_On_Tail" in comp.columns:
        free = (comp["Installed_On_Tail"].isna()
                | comp["Installed_On_Tail"].astype(str).str.strip().isin(["", "nan", "None"]))
        for et, cnt in comp[(comp["Unit_Kind"] == "ENGINE") & free]["Unit_Type"].value_counts().items():
            c = _ENG_CAT.get(str(et))
            if c in sp:
                sp[c] += int(cnt)
        n_apu = int(((comp["Unit_Kind"] == "APU") & free).sum())
    return {"liveTotal": total, "neo": cat["NEO"], "ceo": cat["CEO"], "atr": cat["ATR"],
            "spareNeoEngines": sp["NEO"], "spareCeoEngines": sp["CEO"],
            "spareAtrEngines": sp["ATR"], "spareApus": n_apu}


@app.get("/fleet/aircraft")
def fleet_aircraft():
    ds = _ds()
    return _aircraft_list(ds) if ds is not None else []


@app.get("/fleet/status")
def fleet_status():
    ds = _ds()
    return _status_rows(ds) if ds is not None else []


@app.get("/forecast/events")
def forecast_events():
    ds = _ds()
    return _events(ds) if ds is not None else []


@app.get("/llps")
def llps():
    ds = _ds()
    return _llps(ds) if ds is not None else []


@app.get("/powerplant/units")
def powerplant_units():
    ds = _ds()
    return _power_units(ds) if ds is not None else []


@app.get("/fleet/utilisation")
def fleet_utilisation():
    ds = _ds()
    return _utilisation(ds) if ds is not None else []


@app.get("/data/sheets")
def data_sheets():
    ds = _ds()
    return _sheets(ds) if ds is not None else []


@app.get("/fleet/aircraft/{tail}")
def aircraft_detail(tail: str):
    ds = _ds()
    if ds is None:
        return None
    ac = next((a for a in _aircraft_list(ds) if a["tail"] == tail), None)
    if ac is None:
        raise HTTPException(404, "Unknown tail")
    status = _status_rows(ds)
    power = _power_units(ds)
    evs = _events(ds)
    engines = [u for u in power if u["onTail"] == tail and u["kind"] == "Engine"]
    apus = [u for u in power if u["onTail"] == tail and u["kind"] == "APU"]
    return {
        "aircraft": ac,
        "checks": [s for s in status if s["tail"] == tail and s["kind"] == "Airframe"],
        "engines": engines,
        "apu": apus[0] if apus else None,
        "propellers": [s for s in status if s["tail"] == tail and s["kind"] == "Propeller"],
        "events": [e for e in evs if e["tail"] == tail],
    }


@app.post("/data/upload")
async def data_upload(file: UploadFile = File(...)):
    raw = await file.read()
    if not raw:
        raise HTTPException(400, "Empty file.")
    try:
        ds = load(raw)
        proj = engine.run(ds)
    except Exception as e:
        raise HTTPException(422, f"Could not process workbook: {e}")
    STATE["ds"], STATE["proj"] = ds, proj
    fleet = ds.get("Fleet")
    STATE["airline"] = _str(fleet.iloc[0]["Operator"]) if (fleet is not None and not fleet.empty and "Operator" in fleet.columns) else ""
    return {"date": _snapshot_date(ds), "airline": STATE["airline"], "loaded": True}


@app.put("/data/sheets")
async def put_sheets(sheets: list[dict]):
    try:
        raw = _sheets_to_bytes(sheets)
        ds = load(raw)
        proj = engine.run(ds)
    except Exception as e:
        raise HTTPException(422, f"Could not apply edits: {e}")
    STATE["ds"], STATE["proj"] = ds, proj
    return {"ok": True}


_AI_SYSTEM = (
    "You are a senior fleet maintenance planning analyst for an airline. Answer using ONLY the "
    "data provided below (a snapshot of the current fleet and its projected shop visits). Be "
    "concise and specific: cite numbers and tail/serial IDs, highlight risks such as overdue shop "
    "visits or clustering of events, and give planning-useful observations. If the answer is not "
    "in the data, say so plainly — never invent values."
)


def _ai_context(ds) -> str:
    import json
    rows = _status_rows(ds)
    evs = _events(ds)
    overdue = sorted([r for r in rows if r["remaining"] is not None and r["remaining"] < 0],
                     key=lambda r: r["remaining"])
    upcoming = sorted(evs, key=lambda e: e["dueDate"])[:30]
    lines = [f"SNAPSHOT DATE: {_snapshot_date(ds)}",
             f"FLEET SUMMARY: {json.dumps(summary())}",
             f"COUNTS: {len(rows)} unit×event rows, {len(overdue)} overdue, {len(evs)} projected events."]
    # full aircraft roster with installed engines + APU (answers 'engines of tail X')
    fleet = ds.get("Fleet")
    if fleet is not None and not fleet.empty:
        lines.append("AIRCRAFT ROSTER (tail · type · family · LH engine · RH engine · APU · lessor):")
        for _, r in fleet.iterrows():
            lines.append(f"  {_str(r.get('Tail'))} · {_str(r.get('Aircraft_Family'))} · "
                         f"{_big_cat(r.get('Aircraft_Family'))} · LH:{_str(r.get('LHENG_ESN'))} · "
                         f"RH:{_str(r.get('RHENG_ESN'))} · APU:{_str(r.get('APU_SN'))} · {_str(r.get('Owner'))}")
    # spare engines
    comp = ds.get("Components")
    if comp is not None and not comp.empty and "Installed_On_Tail" in comp.columns:
        free = (comp["Installed_On_Tail"].isna()
                | comp["Installed_On_Tail"].astype(str).str.strip().isin(["", "nan", "None"]))
        sp = comp[(comp["Unit_Kind"] == "ENGINE") & free]
        if not sp.empty:
            lines.append(f"SPARE ENGINES ({len(sp)}): " +
                         ", ".join(f"{_str(r['Component_ID'])}({_str(r['Unit_Type'])})" for _, r in sp.iterrows()))
    lines.append("ALL OVERDUE ITEMS (tail · unit · type · event · remaining · %used):")
    lines += [f"  {r['tail']} · {r['unit']} · {r['type']} · {r['event']} · {r['remaining']} · {r['pctLimit']}%"
              for r in overdue[:40]]
    lines.append("NEXT UPCOMING SHOP VISITS (date · tail · unit · event · reason):")
    lines += [f"  {e['dueDate']} · {e['tail']} · {e['unit']} · {e['event']} · {e['reason']}" for e in upcoming]
    return "\n".join(lines)


@app.post("/ai/ask")
async def ai_ask(body: dict):
    question = str((body or {}).get("question", "")).strip()
    if not question:
        raise HTTPException(400, "Please include a question.")
    ds = _ds()
    if ds is None:
        raise HTTPException(400, "Upload a workbook first.")
    key = os.environ.get("ANTHROPIC_API_KEY")
    if not key:
        raise HTTPException(400, "No ANTHROPIC_API_KEY configured on the server.")
    try:
        import anthropic
        import httpx2
        # Anaconda's bundled decompressor is broken for this SDK's httpx2 stack;
        # request identity (uncompressed) responses to bypass it.
        client = anthropic.Anthropic(
            api_key=key,
            http_client=httpx2.Client(headers={"accept-encoding": "identity"}, timeout=60))
        msg = client.messages.create(
            model="claude-opus-4-8", max_tokens=2000,
            thinking={"type": "adaptive"}, system=_AI_SYSTEM,
            messages=[{"role": "user", "content": f"{_ai_context(ds)}\n\nQUESTION: {question}"}],
        )
        answer = "".join(b.text for b in msg.content if getattr(b, "type", "") == "text")
        return {"answer": answer.strip()}
    except Exception as e:
        raise HTTPException(502, f"AI request failed: {e}")


@app.get("/export/{name}")
def export(name: str):
    ds = _ds()
    if name == "sample":
        if not os.path.exists(_SAMPLE):
            raise HTTPException(404, "Sample not bundled.")
        return StreamingResponse(open(_SAMPLE, "rb"), media_type=_XLSX,
                                 headers={"Content-Disposition": "attachment; filename=SAMPLE_fleet_database.xlsx"})
    if ds is None:
        raise HTTPException(404, "No data loaded. Upload a workbook first.")
    if name == "projection":
        data = build_workbook(STATE["proj"])
        fn = "FLEET_MAINTENANCE_PROJECTION.xlsx"
    elif name == "workbook":
        data = _sheets_to_bytes(_sheets(ds))
        fn = "fleet_database.xlsx"
    else:
        raise HTTPException(404, "Unknown export")
    return StreamingResponse(io.BytesIO(data), media_type=_XLSX,
                             headers={"Content-Disposition": f"attachment; filename={fn}"})
