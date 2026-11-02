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

from fleet_erp.ingestion.loader import load, Dataset  # noqa: E402
from fleet_erp.technical import engine               # noqa: E402
from fleet_erp.export.excel import build_workbook     # noqa: E402

app = FastAPI(title="Fleet Forecast API", version="1.0")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_credentials=False,
                   allow_methods=["*"], allow_headers=["*"])

# single in-memory dataset (Phase 1, no DB). None until a workbook is uploaded.
STATE: dict = {"ds": None, "ds0": None, "proj": None, "airline": ""}
_SAMPLE = os.path.join(_ROOT, "assets", "SAMPLE_fleet_database.xlsx")
_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
# last-uploaded workbook, cached so a backend restart doesn't lose the session (gitignored)
_CACHE = os.path.join(_ROOT, "data", ".session_cache.xlsx")

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


def _as_of_ts(ds):
    """The reference date for 'active fleet' = the snapshot / as-of date from Config."""
    d = pd.to_datetime(_snapshot_date(ds), errors="coerce")
    return d if pd.notna(d) else pd.Timestamp.today().normalize()


def _active_today(fl, ref=None):
    if fl is None or fl.empty:
        return fl
    d = fl.copy()
    for c in ("Delivery_Date", "Redelivery_Date", "Extension_Date"):
        if c in d.columns:
            d[c] = pd.to_datetime(d[c], errors="coerce")
    today = ref if ref is not None else pd.Timestamp.today().normalize()
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
def _aircraft_list(ds, ref=None) -> list[dict]:
    fleet = ds.get("Fleet")
    if fleet is None or fleet.empty:
        return []
    util = _util_by_tail(ds)
    active_tails = set(_active_today(fleet, ref if ref is not None else _as_of_ts(ds))["Tail"].astype(str))
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
            "active": tail in active_tails,
            "fhYear": fh, "fcYear": fc,
        })
    return out


def _clean_type(t) -> str:
    """Landing-gear units are stored as '<family> LG' internally to separate their
    events; show just the aircraft type in the UI (kind=Landing Gear distinguishes)."""
    s = _str(t)
    return s[:-3].rstrip() if s.endswith(" LG") else s


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
            "kind": _kind(kind), "type": _clean_type(utype), "pos": _POSL.get(pos, pos),
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
            "reasonText": _str(e.get("Reason")),   # full descriptive wording from the engine
            "position": {"AIRFRAME": "AIRFRAME", "LHENG": "LH", "RHENG": "RH", "APU": "APU",
                         "LHPROP": "LH", "RHPROP": "RH", "LG": "LG", "SPARE": "SPARE"}
                        .get(_str(e.get("Position")), _str(e.get("Position"))),
        })
    return out


def _llps(ds) -> list[dict]:
    reg = ds.get("LLP_Registry")
    comp = ds.get("Components")
    proj = STATE["proj"]
    if reg is None or reg.empty:
        return []
    etype = {}
    if comp is not None and not comp.empty:
        etype = dict(zip(comp["Component_ID"].astype(str), comp["Unit_Type"].astype(str)))
    # projected SV date + reason per LLP come from the PROJECTION (same as Streamlit),
    # keyed by (engine serial, LLP SN) via the LLP_<SN> event id — NOT the file's Expected_Date.
    llp_proj = {}
    if proj is not None and proj.events is not None and not proj.events.empty:
        pe = proj.events[proj.events["Trigger"] == "LLP_REPLACEMENT"].sort_values("Date")
        for _, e in pe.iterrows():
            eid = str(e.get("Event_ID"))
            sn = eid[4:] if eid.startswith("LLP_") else eid
            key = (str(e.get("Unit_ID")), sn)
            if key not in llp_proj:   # keep the SOONEST (next-due) replacement, not later cycles
                llp_proj[key] = (_date(e.get("Date")), _reason(e.get("Trigger")), _str(e.get("Reason")))
    out = []
    for i, r in reg.reset_index(drop=True).iterrows():
        uid = _str(r.get("Unit_ID"))
        sn = _str(r.get("LLP_SN"))
        lim = _num(r.get("Limit_Cycles"))
        togo = _num(r.get("ToGo"))
        pdate, prs, prtext = llp_proj.get((uid, sn), ("", "llp", ""))
        out.append({
            "id": str(i), "unit": uid, "engineType": etype.get(uid, ""),
            "part": _str(r.get("Description")), "pn": _str(r.get("LLP_PN")),
            "sn": sn, "limit": lim,
            "csn": (lim - togo) if (isinstance(lim, (int, float)) and isinstance(togo, (int, float))) else _num(r.get("SLSV")),
            "remaining": togo,
            "projectedDate": pdate,                 # from the projection (correct), not Expected_Date
            "reason": prs, "reasonText": prtext or "Reached its cycle limit",
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


_SKIP_SHEETS = {"Instructions", "_Lists"}


def _sheets(ds) -> list[dict]:
    out = []
    for name, df in ds.sheets.items():
        if df is None or name in _SKIP_SHEETS:
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
    active = _active_today(ds.get("Fleet"), _as_of_ts(ds))
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
def fleet_aircraft(as_of: str | None = None):
    ds = _ds()
    if ds is None:
        return []
    ref = pd.to_datetime(as_of, errors="coerce") if as_of else None
    return _aircraft_list(ds, ref if (ref is not None and not pd.isna(ref)) else None)


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


@app.get("/fleet/utilisation/units")
def utilisation_units():
    """Per-unit utilisation assumptions (every tail / engine / APU / propeller) for drill-down."""
    ds = _ds()
    if ds is None:
        return []
    util = ds.get("Utilization")
    if util is None or util.empty:
        return []
    comp = ds.get("Components")
    fleet = ds.get("Fleet")
    ctype = dict(zip(comp["Component_ID"].astype(str), comp["Unit_Type"].astype(str))) \
        if comp is not None and not comp.empty else {}
    ftype = dict(zip(fleet["Tail"].astype(str), fleet["Aircraft_Family"].astype(str))) \
        if fleet is not None and not fleet.empty else {}
    tails = {str(t) for t in fleet["Tail"]} if fleet is not None and not fleet.empty else set()
    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    out = []
    for _, r in util.iterrows():
        uid = _str(r.get("Unit_ID"))
        out.append({
            "unitId": uid, "kind": _kind(r.get("Unit_Kind")),
            "type": _clean_type(ctype.get(uid) or ftype.get(uid) or ""),
            "onTail": uid if uid in tails else _str(comp.loc[comp["Component_ID"].astype(str) == uid, "Installed_On_Tail"].iloc[0]) if (comp is not None and not comp.empty and (comp["Component_ID"].astype(str) == uid).any()) else "",
            "annualFh": _num(r.get("Annual_FH")), "annualFc": _num(r.get("Annual_FC")),
            "ratio": _num(r.get("FH_FC_Ratio")), "evolutionPct": _num(r.get("Annual_Evolution_Pct")),
            "minFh": _num(r.get("Min_Annual_FH")), "maxFh": _num(r.get("Max_Annual_FH")),
            "aogLeadInMonths": _num(r.get("AOG_Lead_In_Months")),
            "monthly": [_num(r.get(m)) for m in months],
        })
    return out


@app.get("/data/sheets")
def data_sheets():
    ds = _ds()
    return _sheets(ds) if ds is not None else []


@app.get("/data/sheets/names")
def data_sheet_names():
    ds = _ds()
    return [s["name"] for s in _sheets(ds)] if ds is not None else []


@app.get("/data/sheets/{name}")
def data_sheet_one(name: str):
    ds = _ds()
    if ds is None:
        raise HTTPException(404, "No data loaded.")
    for s in _sheets(ds):
        if s["name"] == name:
            return s
    raise HTTPException(404, "Unknown sheet")


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
    STATE["ds"], STATE["ds0"], STATE["proj"] = ds, ds, proj
    fleet = ds.get("Fleet")
    STATE["airline"] = _str(fleet.iloc[0]["Operator"]) if (fleet is not None and not fleet.empty and "Operator" in fleet.columns) else ""
    try:   # cache so a backend restart restores the session
        with open(_CACHE, "wb") as _c:
            _c.write(raw)
    except OSError:
        pass
    return {"date": _snapshot_date(ds), "airline": STATE["airline"], "loaded": True}


def _restore_session():
    """On startup, reload the last-uploaded workbook so a restart keeps the data."""
    if STATE["ds"] is not None or not os.path.exists(_CACHE):
        return
    try:
        raw = open(_CACHE, "rb").read()
        ds = load(raw)
        STATE["ds"], STATE["ds0"], STATE["proj"] = ds, ds, engine.run(ds)
        fleet = ds.get("Fleet")
        STATE["airline"] = _str(fleet.iloc[0]["Operator"]) if (fleet is not None and not fleet.empty and "Operator" in fleet.columns) else ""
    except Exception:
        pass


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
    "You are a senior fleet maintenance planning analyst for an airline. A summary of the current "
    "fleet is provided below, and you ALSO have tools to query the full projection and LLP data on "
    "demand — use them whenever a question needs detail beyond the summary. Tools: unit_events "
    "(all projected events for an engine/APU/landing-gear/propeller/tail), llps (life-limited parts "
    "for an engine, sorted soonest-first — use this to find which LLP hits first), search_events "
    "(filter projected events), unit_status (current readings for a unit or tail). Always call a "
    "tool rather than saying data is missing. Be concise and specific: cite numbers, tail/serial "
    "IDs and dates; flag risks. Never invent values — if a tool returns nothing, say so."
)


# ── AI tools: let the analyst query the FULL uploaded data + projection ──
def _tool_unit_events(unit_id: str):
    proj = STATE["proj"]
    if proj is None or proj.events is None or proj.events.empty:
        return []
    ev = proj.events
    m = ev[ev["Unit_ID"].astype(str).str.upper() == str(unit_id).upper()]
    return [{"date": _date(r["Date"]), "event": _str(r["Event"]),
             "driver": _str(r["Driver"]), "reason": _str(r["Reason"])}
            for _, r in m.sort_values("Date").head(120).iterrows()]


def _tool_llps(unit_id: str):
    reg = STATE["ds"].get("LLP_Registry")
    if reg is None or reg.empty:
        return []
    m = reg[reg["Unit_ID"].astype(str).str.upper() == str(unit_id).upper()].copy()
    if m.empty:
        return []
    m["_togo"] = pd.to_numeric(m.get("ToGo"), errors="coerce")
    m = m.sort_values("_togo", na_position="last")
    return [{"part": _str(r.get("Description")), "pn": _str(r.get("LLP_PN")),
             "sn": _str(r.get("LLP_SN")), "limit_cycles": _num(r.get("Limit_Cycles")),
             "cycles_to_go": _num(r.get("ToGo")), "slsv": _num(r.get("SLSV")),
             "expected_date": _date(r.get("Expected_Date"))}
            for _, r in m.head(80).iterrows()]


def _tool_search_events(kind=None, unit_type=None, event=None, tail=None, unit=None, limit=60):
    evs = _events(STATE["ds"])

    def ok(e):
        return ((not kind or e["kind"].lower() == str(kind).lower())
                and (not unit_type or str(unit_type).lower() in e["type"].lower())
                and (not event or str(event).lower() in e["event"].lower())
                and (not tail or e["tail"].upper() == str(tail).upper())
                and (not unit or e["unit"].upper() == str(unit).upper()))
    return sorted([e for e in evs if ok(e)], key=lambda e: e["dueDate"])[: int(limit or 60)]


def _tool_unit_status(unit_or_tail: str):
    u = str(unit_or_tail).upper()
    return [r for r in _status_rows(STATE["ds"]) if r["unit"].upper() == u or r["tail"].upper() == u][:80]


_TOOLS = [
    {"name": "unit_events", "description": "All projected shop-visit/LLP events for one unit (engine serial, APU serial, landing-gear id, propeller id, or aircraft tail), sorted by date.",
     "input_schema": {"type": "object", "properties": {"unit_id": {"type": "string"}}, "required": ["unit_id"]}},
    {"name": "llps", "description": "Life-limited parts for an engine serial, sorted soonest-first (smallest cycles_to_go). Use to find which LLP hits first.",
     "input_schema": {"type": "object", "properties": {"unit_id": {"type": "string"}}, "required": ["unit_id"]}},
    {"name": "search_events", "description": "Search projected events with optional filters (kind, unit_type, event, tail, unit).",
     "input_schema": {"type": "object", "properties": {"kind": {"type": "string"}, "unit_type": {"type": "string"}, "event": {"type": "string"}, "tail": {"type": "string"}, "unit": {"type": "string"}, "limit": {"type": "integer"}}}},
    {"name": "unit_status", "description": "Current status (since-last, limit, remaining, TSN/CSN) for a unit serial or an aircraft tail.",
     "input_schema": {"type": "object", "properties": {"unit_or_tail": {"type": "string"}}, "required": ["unit_or_tail"]}},
]


def _dispatch(name, inp):
    try:
        inp = inp or {}
        if name == "unit_events":
            return _tool_unit_events(inp.get("unit_id", ""))
        if name == "llps":
            return _tool_llps(inp.get("unit_id", ""))
        if name == "search_events":
            return _tool_search_events(**inp)
        if name == "unit_status":
            return _tool_unit_status(inp.get("unit_or_tail", ""))
    except Exception as e:
        return {"error": str(e)}
    return {"error": "unknown tool"}


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
        import json as _json
        import anthropic
        import httpx2
        # Anaconda's bundled decompressor is broken for this SDK's httpx2 stack;
        # request identity (uncompressed) responses to bypass it.
        client = anthropic.Anthropic(
            api_key=key,
            http_client=httpx2.Client(headers={"accept-encoding": "identity"}, timeout=90))
        messages = [{"role": "user", "content": f"{_ai_context(ds)}\n\nQUESTION: {question}"}]
        answer = ""
        for _ in range(6):   # tool-use loop
            msg = client.messages.create(
                model="claude-opus-4-8", max_tokens=3000,
                thinking={"type": "adaptive"}, system=_AI_SYSTEM,
                tools=_TOOLS, messages=messages)
            if msg.stop_reason == "tool_use":
                messages.append({"role": "assistant", "content": msg.content})
                results = []
                for b in msg.content:
                    if getattr(b, "type", "") == "tool_use":
                        out = _dispatch(b.name, b.input)
                        results.append({"type": "tool_result", "tool_use_id": b.id,
                                        "content": _json.dumps(out)[:12000]})
                messages.append({"role": "user", "content": results})
                continue
            answer = "".join(b.text for b in msg.content if getattr(b, "type", "") == "text")
            break
        return {"answer": answer.strip() or "No answer produced."}
    except Exception as e:
        raise HTTPException(502, f"AI request failed: {e}")


@app.post("/report")
def fleet_report():
    """Claude-generated executive Fleet Update Report (markdown) from the loaded data."""
    ds = _ds()
    if ds is None:
        raise HTTPException(400, "Upload a workbook first.")
    key = os.environ.get("ANTHROPIC_API_KEY")
    if not key:
        raise HTTPException(400, "No ANTHROPIC_API_KEY configured on the server.")
    try:
        import anthropic
        import httpx2
        client = anthropic.Anthropic(
            api_key=key,
            http_client=httpx2.Client(headers={"accept-encoding": "identity"}, timeout=120))
        instr = (
            "Write an executive FLEET MAINTENANCE UPDATE REPORT in clean markdown for a McKinsey-style "
            "briefing. Use '## ' section headings and short bullet lists. Sections: Executive Summary; "
            "Fleet Composition & Ownership; Near-Term Maintenance Outlook (next 12 months — due shop "
            "visits and clustering); Overdue & At-Risk Items (prioritised, with tail/serial and %used); "
            "Spare-Engine Coverage & Risk; Recommendations. Cite concrete numbers and IDs; be concise, "
            "insight-led and board-ready. Do not invent values.")
        msg = client.messages.create(
            model="claude-opus-4-8", max_tokens=4000, thinking={"type": "adaptive"},
            system=_AI_SYSTEM, tools=_TOOLS,
            messages=[{"role": "user", "content": f"{_ai_context(ds)}\n\n{instr}"}])
        import json as _json
        messages = [{"role": "user", "content": f"{_ai_context(ds)}\n\n{instr}"}]
        for _ in range(5):
            if msg.stop_reason == "tool_use":
                messages.append({"role": "assistant", "content": msg.content})
                results = [{"type": "tool_result", "tool_use_id": b.id,
                            "content": _json.dumps(_dispatch(b.name, b.input))[:12000]}
                           for b in msg.content if getattr(b, "type", "") == "tool_use"]
                messages.append({"role": "user", "content": results})
                msg = client.messages.create(
                    model="claude-opus-4-8", max_tokens=4000, thinking={"type": "adaptive"},
                    system=_AI_SYSTEM, tools=_TOOLS, messages=messages)
                continue
            break
        report = "".join(b.text for b in msg.content if getattr(b, "type", "") == "text")
        return {"report": report.strip(), "snapshot": _snapshot_date(ds), "airline": STATE["airline"]}
    except Exception as e:
        raise HTTPException(502, f"Report failed: {e}")


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


# restore the last session on startup (survives restarts)
_restore_session()
