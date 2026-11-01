"""
Fleet Forecast — REST API (Phase 1, no database).

A thin FastAPI layer over the proven Python engine in ``src/fleet_erp``.
Upload a workbook → it is parsed and projected → results are served as JSON
for the Lovable React UI. State lives in memory keyed by a session id.

IMPORTANT: nothing about the fleet is hardcoded. Unit kinds, types, events,
drivers and limits are all read from the uploaded file — the file dictates
everything, exactly like the Streamlit app.
"""
from __future__ import annotations

import io
import os
import sys
import uuid

import numpy as np
import pandas as pd
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse

# make src/ importable
_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(_ROOT, "src"))

from fleet_erp.ingestion.loader import load          # noqa: E402
from fleet_erp.technical import engine               # noqa: E402
from fleet_erp.export.excel import build_workbook     # noqa: E402

app = FastAPI(title="Fleet Forecast API", version="1.0")

# Dev CORS — allow the local Vite dev server and Lovable/Vercel previews.
# Tighten allow_origins to your real UI domain(s) before production.
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

# in-memory session store: {session_id: {"ds":..., "proj":...}}
SESSIONS: dict[str, dict] = {}
_SAMPLE = os.path.join(_ROOT, "assets", "SAMPLE_fleet_database.xlsx")


# ───────────────────────── helpers ─────────────────────────
def _num(v):
    """JSON-safe number or None."""
    try:
        f = float(v)
    except (TypeError, ValueError):
        return None
    if pd.isna(f):
        return None
    return int(f) if f == int(f) else round(f, 2)


def _str(v):
    if v is None:
        return None
    try:
        if pd.isna(v):
            return None
    except (TypeError, ValueError):
        pass
    s = str(v).strip()
    return s or None


def _date(v):
    try:
        if pd.isna(v):
            return None
    except (TypeError, ValueError):
        pass
    try:
        return pd.to_datetime(v).strftime("%Y-%m-%d")
    except (TypeError, ValueError):
        return _str(v)


def _records(df: pd.DataFrame) -> list[dict]:
    """Serialize an arbitrary DataFrame to JSON-safe records."""
    if df is None or df.empty:
        return []
    out = []
    for _, r in df.iterrows():
        row = {}
        for c in df.columns:
            v = r[c]
            if isinstance(v, (pd.Timestamp,)):
                row[str(c)] = _date(v)
            elif isinstance(v, (int, float, np.integer, np.floating)):
                row[str(c)] = _num(v)
            else:
                row[str(c)] = _str(v)
        out.append(row)
    return out


def _session(sid: str) -> dict:
    s = SESSIONS.get(sid)
    if not s:
        raise HTTPException(status_code=404, detail="Unknown or expired session. Re-upload the workbook.")
    return s


def _active_today(fl: pd.DataFrame) -> pd.DataFrame:
    """Aircraft in service today: Delivery ≤ today ≤ Redelivery (or Extension)."""
    if fl is None or fl.empty:
        return fl
    d = fl.copy()
    for c in ("Delivery_Date", "Redelivery_Date", "Extension_Date"):
        if c in d.columns:
            d[c] = pd.to_datetime(d[c], errors="coerce")
    today = pd.Timestamp.today().normalize()
    end = None
    if "Redelivery_Date" in d.columns:
        end = d["Redelivery_Date"]
    if "Extension_Date" in d.columns:
        end = d["Extension_Date"].fillna(end) if end is not None else d["Extension_Date"]
    m = pd.Series(True, index=d.index)
    if "Delivery_Date" in d.columns:
        m &= d["Delivery_Date"].isna() | (d["Delivery_Date"] <= today)
    if end is not None:
        m &= end.isna() | (end >= today)
    return d[m]


def _big_cat(name: str) -> str:
    f = str(name).upper()
    if "NEO" in f:
        return "NEO"
    if "CEO" in f:
        return "CEO"
    if "ATR" in f or "DASH" in f or "DH" in f or "Q400" in f:
        return "ATR"
    return "Other"


# engine-model → aircraft category is derived from the engine TYPE string in the
# file; this is a grouping heuristic for the KPI cards, not a fleet definition.
_ENG_CAT = {"PW1100": "NEO", "PW1500": "NEO", "LEAP": "NEO",
            "V2500": "CEO", "CFM56": "CEO", "PW127": "ATR"}


def _snapshot_date(ds) -> str | None:
    cfg = ds.get("Config")
    if cfg is not None and not cfg.empty and "Setting" in cfg.columns:
        m = cfg[cfg["Setting"].astype(str) == "As_Of_Date"]
        if not m.empty:
            return _date(m.iloc[0]["Value"])
    return None


# ───────────────────────── routes ─────────────────────────
@app.get("/")
def root():
    return {"service": "Fleet Forecast API", "status": "ok",
            "sessions": len(SESSIONS),
            "docs": "/docs", "sample": "/api/sample"}


@app.get("/api/sample")
def sample():
    if not os.path.exists(_SAMPLE):
        raise HTTPException(404, "Sample workbook not bundled.")
    return StreamingResponse(
        open(_SAMPLE, "rb"),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=SAMPLE_fleet_database.xlsx"})


@app.post("/api/session/upload")
async def upload(file: UploadFile = File(...)):
    raw = await file.read()
    if not raw:
        raise HTTPException(400, "Empty file.")
    try:
        ds = load(raw)
        proj = engine.run(ds)
    except Exception as e:  # surface parse/projection errors cleanly
        raise HTTPException(422, f"Could not process workbook: {e}")
    sid = uuid.uuid4().hex
    SESSIONS[sid] = {"ds": ds, "proj": proj}
    return {"session_id": sid,
            "snapshot_date": _snapshot_date(ds),
            "warnings": list(ds.warnings)}


@app.get("/api/{sid}/overview")
def overview(sid: str):
    ds = _session(sid)["ds"]
    fleet = ds.get("Fleet")
    active = _active_today(fleet)
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

    return {"snapshot_date": _snapshot_date(ds),
            "fleet": {"total": total, "neo": cat["NEO"], "ceo": cat["CEO"], "atr": cat["ATR"]},
            "spares": {"neo_eng": sp["NEO"], "ceo_eng": sp["CEO"], "atr_eng": sp["ATR"], "apu": n_apu}}


_POSL = {"LHENG": "LH", "RHENG": "RH", "APU": "APU", "LHPROP": "LH", "RHPROP": "RH",
         "AIRFRAME": "—", "SPARE": "Spare"}


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
            "tail": _str(tail), "unit": uid, "kind": kind, "type": utype,
            "pos": _POSL.get(pos, pos), "event": evid, "driver": drv,
            "sinceLastEvent": _num(consumed), "cyclesSinceSV": _num(cyc),
            "limit": _num(hln), "remaining": _num(remaining), "pctOfLimit": _num(pct),
            "tsnAtLsv": _num(tsn_l), "csnAtLsv": _num(csn_l),
            "tsnNow": _num(tsn), "csnNow": _num(csn), "lastSv": _date(r.get("Last_SV_Date")),
        })
    rows.sort(key=lambda x: (x["pctOfLimit"] is None, -(x["pctOfLimit"] or 0)))
    return rows


@app.get("/api/{sid}/fleet-status")
def fleet_status(sid: str, kind: str | None = None, type: str | None = None,
                 event: str | None = None, tail: str | None = None):
    ds = _session(sid)["ds"]
    rows = _status_rows(ds)
    # filter options are derived from the DATA, never a hardcoded list
    filters = {
        "kinds": sorted({r["kind"] for r in rows if r["kind"]}),
        "types": sorted({r["type"] for r in rows if r["type"]}),
        "events": sorted({r["event"] for r in rows if r["event"]}),
        "tails": sorted({r["tail"] for r in rows if r["tail"]} | {r["unit"] for r in rows if r["unit"]}),
    }

    def pick(v):
        return [s.strip() for s in v.split(",")] if v else None
    fk, ft, fe, fl = pick(kind), pick(type), pick(event), pick(tail)
    if fk:
        rows = [r for r in rows if r["kind"] in fk]
    if ft:
        rows = [r for r in rows if r["type"] in ft]
    if fe:
        rows = [r for r in rows if r["event"] in fe]
    if fl:
        rows = [r for r in rows if r["tail"] in fl or r["unit"] in fl]
    return {"rows": rows, "filters": filters}


@app.get("/api/{sid}/powerplant")
def powerplant(sid: str):
    ds = _session(sid)["ds"]
    comp = ds.get("Components")
    us = ds.get("Unit_Status")
    llp = ds.get("LLP_Registry")
    if comp is None or comp.empty:
        return {"kpis": {}, "units": []}
    pw = comp[comp["Unit_Kind"].isin(["ENGINE", "APU"])].copy()
    posmap = {"LHENG": "LH", "RHENG": "RH", "APU": "APU", "SPARE": "Spare"}
    agg = {}
    if us is not None and not us.empty:
        acols = {"tsn": ("TSN", "first"), "csn": ("CSN", "first"),
                 "lastSvDate": ("Last_SV_Date", "first")}
        if "TSN_At_LSV" in us.columns:
            acols["tsnAtLsv"] = ("TSN_At_LSV", "first")
        if "WS_Type" in us.columns:
            acols["workscope"] = ("WS_Type", "first")
        g = us.groupby(us["Unit_ID"].astype(str)).agg(**acols)
        agg = g.to_dict("index")
    lc = {}
    if llp is not None and not llp.empty:
        lc = llp.groupby(llp["Unit_ID"].astype(str)).size().to_dict()
    units = []
    for _, r in pw.iterrows():
        cid = str(r["Component_ID"])
        pos = str(r.get("Position") or "")
        a = agg.get(cid, {})
        tail = _str(r.get("Installed_On_Tail"))
        units.append({
            "serial": cid, "type": _str(r.get("Unit_Type")), "kind": str(r["Unit_Kind"]),
            "pos": posmap.get(pos, pos),
            "state": "Spare" if (pos == "SPARE" or not tail) else "Installed",
            "onTail": tail,
            "tsn": _num(a.get("tsn")), "tsnAtLsv": _num(a.get("tsnAtLsv")),
            "csn": _num(a.get("csn")), "lastSvDate": _date(a.get("lastSvDate")),
            "workscope": _str(a.get("workscope")), "llps": int(lc.get(cid, 0)),
        })
    eng = [u for u in units if u["kind"] == "ENGINE"]
    apu = [u for u in units if u["kind"] == "APU"]
    kpis = {
        "installedEngines": sum(u["state"] == "Installed" for u in eng),
        "spareEngines": sum(u["state"] == "Spare" for u in eng),
        "installedApus": sum(u["state"] == "Installed" for u in apu),
        "spareApus": sum(u["state"] == "Spare" for u in apu),
    }
    return {"kpis": kpis, "units": units}


@app.get("/api/{sid}/llp")
def llp(sid: str):
    ds = _session(sid)["ds"]
    reg = ds.get("LLP_Registry")
    if reg is None or reg.empty:
        return {"rows": []}
    rows = []
    for _, r in reg.iterrows():
        rows.append({
            "unitId": _str(r.get("Unit_ID")), "pn": _str(r.get("LLP_PN")),
            "sn": _str(r.get("LLP_SN")), "description": _str(r.get("Description")),
            "limitCycles": _num(r.get("Limit_Cycles")), "slsv": _num(r.get("SLSV")),
            "toGo": _num(r.get("ToGo")), "projectedSv": _date(r.get("Expected_Date")),
            "reason": _str(r.get("Reason")) or "",
        })
    return {"rows": rows}


@app.get("/api/{sid}/projection")
def projection(sid: str):
    proj = _session(sid)["proj"]
    return {"events": _records(proj.events)}


@app.get("/api/{sid}/calendar")
def calendar(sid: str):
    proj = _session(sid)["proj"]
    return {"months": _records(proj.sv_plan_v1)}


@app.get("/api/{sid}/aircraft/{tail}")
def aircraft(sid: str, tail: str):
    s = _session(sid)
    ds, proj = s["ds"], s["proj"]
    fleet = ds.get("Fleet")
    row = None
    if fleet is not None and not fleet.empty:
        m = fleet[fleet["Tail"].astype(str) == tail]
        if not m.empty:
            row = _records(m)[0]
    events = proj.events
    ev = events[events.get("Tail", pd.Series(dtype=str)).astype(str) == tail] if events is not None and not events.empty and "Tail" in events.columns else None
    comp = ds.get("Components")
    units = comp[comp["Installed_On_Tail"].astype(str) == tail] if comp is not None and not comp.empty else None
    return {"tail": tail, "aircraft": row,
            "events": _records(ev) if ev is not None else [],
            "units": _records(units) if units is not None else []}


@app.get("/api/{sid}/export")
def export(sid: str):
    proj = _session(sid)["proj"]
    data = build_workbook(proj)
    return StreamingResponse(
        io.BytesIO(data),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=FLEET_MAINTENANCE_PROJECTION.xlsx"})
