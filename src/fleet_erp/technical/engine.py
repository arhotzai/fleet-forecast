"""
Technical Projection Engine (Stage 1)
=====================================
Projects every serialized unit (airframe / engine / APU) month-by-month from the
As-Of date, accumulating each monitored event's "since last shop visit" (SLSV)
against its driver (FH / FC / calendar), triggering a reset when it reaches its
hard limit, and cascading HARD-linked child resets. LLPs are projected from the
LLP report's SLSV (= life_limit − togo).

Output: a tidy Event Log DataFrame + the final per-unit status. No wide-column
string parsing — everything is driven by the relational tables.
"""

from __future__ import annotations

from dataclasses import dataclass

import numpy as np
import pandas as pd
from dateutil.relativedelta import relativedelta

from fleet_erp.ingestion.loader import Dataset

DEFAULT_HORIZON_MONTHS = 144


@dataclass
class Projection:
    events: pd.DataFrame          # one row per shop-visit / replacement event
    unit_status: pd.DataFrame     # final SLSV per unit × event
    dates: list[pd.Timestamp]
    horizon_end: pd.Timestamp
    sv_plan_v1: pd.DataFrame = None     # month × aircraft calendar of SV labels
    sv_plan_v2: pd.DataFrame = None     # event × aircraft matrix of dates
    aircraft_frames: dict = None        # {tail: monthly counter DataFrame}


def _sv_label(position: str, event: str, is_llp: bool, unit_id: str | None = None) -> str:
    """Readable shop-visit label for the plans.

    Engine/APU events are prefixed with the actual serial (ESN / APU serial),
    not the position (LHENG/RHENG/APU), so each cell names the real component.
    """
    base = "LLP" if is_llp else event
    if position in ("AIRFRAME", "SPARE"):
        return f"{base}-SV"
    prefix = str(unit_id) if unit_id else position
    return f"{prefix} {base}-SV"


# ─────────────────────────────────────────────────────────────
# Configuration helpers
# ─────────────────────────────────────────────────────────────

def _config(ds: Dataset) -> dict:
    cfg = ds.get("Config")
    out = {}
    if not cfg.empty and {"Setting", "Value"} <= set(cfg.columns):
        out = dict(zip(cfg["Setting"].astype(str), cfg["Value"]))
    return out


def _as_of(ds: Dataset) -> pd.Timestamp:
    cfg = _config(ds)
    d = pd.to_datetime(cfg.get("As_Of_Date"), errors="coerce")
    if pd.isna(d):
        d = pd.Timestamp.today()
    return (d + pd.offsets.MonthEnd(0)).normalize()


def _num(v, default=np.nan) -> float:
    try:
        f = float(v)
        return f if not pd.isna(f) else default
    except (TypeError, ValueError):
        return default


# ─────────────────────────────────────────────────────────────
# Utilisation
# ─────────────────────────────────────────────────────────────

class Utilisation:
    """Monthly FH / FC for one unit across the date index."""

    def __init__(self, dates, fh: np.ndarray, fc: np.ndarray):
        self.dates = dates
        self.fh = fh
        self.fc = fc


_MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
           "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def _build_utilisation(ds: Dataset, dates, as_of) -> dict[str, Utilisation]:
    """
    Utilisation = ANNUAL FH/FC distributed across the year by monthly percentages.
    Monthly FH = (Annual_FH grown by evolution, clamped to Min/Max) × month-percentage.
    """
    util = ds.get("Utilization")
    result: dict[str, Utilisation] = {}
    if util.empty or "Unit_ID" not in util.columns:
        return result

    n = len(dates)
    month_idx = np.array([d.month - 1 for d in dates])          # 0..11
    years = np.array([(d.year - as_of.year) + (d.month - as_of.month) / 12.0 for d in dates])
    has_months = all(m in util.columns for m in _MONTHS)

    for _, r in util.iterrows():
        uid = str(r.get("Unit_ID"))
        annual_fh = _num(r.get("Annual_FH"), 0.0)
        annual_fc = _num(r.get("Annual_FC"), 0.0)
        evo = _num(r.get("Annual_Evolution_Pct"), 0.0)
        min_a = _num(r.get("Min_Annual_FH"))
        max_a = _num(r.get("Max_Annual_FH"))
        aog = int(_num(r.get("AOG_Lead_In_Months"), 0))

        # monthly percentages (fall back to even 1/12 if absent)
        if has_months:
            pct = np.array([_num(r.get(m), 0.0) for m in _MONTHS])
            if np.nansum(pct) <= 0:
                pct = np.full(12, 1.0 / 12.0)
        else:
            pct = np.full(12, 1.0 / 12.0)

        growth = np.power(1.0 + evo, years)
        annual_fh_t = annual_fh * growth
        annual_fc_t = annual_fc * growth
        if not np.isnan(min_a):
            annual_fh_t = np.maximum(annual_fh_t, min_a)
        if not np.isnan(max_a):
            annual_fh_t = np.minimum(annual_fh_t, max_a)

        fh = annual_fh_t * pct[month_idx]
        fc = annual_fc_t * pct[month_idx]
        if aog > 0:
            fh[:min(aog, n)] = 0.0
            fc[:min(aog, n)] = 0.0
        result[uid] = Utilisation(dates, fh, fc)
    return result


# ─────────────────────────────────────────────────────────────
# Event definitions
# ─────────────────────────────────────────────────────────────

def _events_by_type(ds: Dataset) -> dict[str, list[dict]]:
    """Return {(Unit_Type, Unit_Kind): [event dicts]} from the single Events sheet."""
    out: dict[str, list[dict]] = {}
    df = ds.get("Events")
    if df.empty:
        return out
    for _, r in df.iterrows():
        key = (str(r.get("Unit_Type")), str(r.get("Unit_Kind")))
        out.setdefault(key, []).append({
            "event_id": str(r.get("Event_ID")),
            "event_name": str(r.get("Event_Name") or r.get("Event_ID")),
            "driver": str(r.get("Driver") or "").upper(),
            "hard_limit": _num(r.get("Hard_Limit")),
            # soft-bundle threshold expressed in YEARS of remaining life
            "soft_years": _num(r.get("Soft_Threshold_Years")),
            # min life (YEARS) that must remain at lease redelivery; if less than
            # this remains at hand-back, a redelivery-condition SV is triggered.
            # Blank = no redelivery requirement for this event.
            "redeliv_min": _num(r.get("Redelivery_Min_Life_Years")),
        })
    return out


def _links_by_type(ds: Dataset) -> dict[str, list[tuple[str, str, str]]]:
    """{Unit_Type: [(source_event, target_event, link_type)]} from Event_Links."""
    out: dict[str, list[tuple[str, str, str]]] = {}
    df = ds.get("Event_Links")
    if df.empty:
        return out
    for _, r in df.iterrows():
        ut = str(r.get("Unit_Type"))
        src = str(r.get("Source_Event") or "").strip()
        tgt = str(r.get("Target_Event") or "").strip()
        lt = str(r.get("Link_Type") or "HARD").upper()
        if src and tgt:
            out.setdefault(ut, []).append((src, tgt, lt))
    return out


def _llp_rules(ds: Dataset) -> dict[str, tuple[str, float]]:
    """{Unit_Type: (llp_link_to_event, llp_soft_threshold_years)} from Unit_Types."""
    out: dict[str, tuple[str, float]] = {}
    df = ds.get("Unit_Types")
    if df.empty:
        return out
    for _, r in df.iterrows():
        ut = str(r.get("Unit_Type"))
        link = r.get("LLP_Link_To_Event")
        yrs = r.get("LLP_Soft_Threshold_Years")
        if yrs is None or (isinstance(yrs, float) and pd.isna(yrs)):
            yrs = r.get("LLP_Soft_Threshold")   # tolerate old column name
        if pd.notna(link) and str(link).strip():
            out[ut] = (str(link).strip(), _num(yrs))
    return out


def _initial_status(ds: Dataset) -> dict[tuple[str, str], dict]:
    """{(Unit_ID, Event_ID): {fh, fc, months}} from Unit_Status."""
    out: dict[tuple[str, str], dict] = {}
    us = ds.get("Unit_Status")
    if us.empty:
        return out
    for _, r in us.iterrows():
        key = (str(r.get("Unit_ID")), str(r.get("Event_ID")))
        # FC distance since the last SV is DERIVED from the counters (CSN now −
        # CSN at last SV), not stored — the real driver for HSI/CSI/RGB is FH and
        # the FC value is only a counter. Fall back to SLSV_FC if counters absent.
        csn, csn_lsv = _num(r.get("CSN")), _num(r.get("CSN_At_LSV"))
        if not np.isnan(csn) and not np.isnan(csn_lsv):
            fc = max(0.0, csn - csn_lsv)
        else:
            fc = _num(r.get("SLSV_FC"), 0.0)
        out[key] = {
            "fh": _num(r.get("SLSV_FH"), 0.0),
            "fc": fc,
            "months": _num(r.get("SLSV_Months"), 0.0),
            # per-unit redelivery min-life (YEARS). Overrides the per-type Events value
            # when present — a lease's return condition can differ per installed unit.
            "redeliv_min": _num(r.get("Redelivery_Min_Life_Years")),
        }
    return out


# ─────────────────────────────────────────────────────────────
# Per-unit projection
# ─────────────────────────────────────────────────────────────

def _within_soft(e, annual_fh_i: float, annual_fc_i: float) -> bool:
    """
    Bundle if remaining life is less than (threshold_years × annual utilisation).
    threshold_years = 1 → within one year of being due; 0.5 → within half a year.
    """
    yrs, lim, slsv, d = e["soft_years"], e["hard_limit"], e["slsv"], e["driver"]
    if np.isnan(yrs) or np.isnan(lim) or lim <= 0:
        return False
    if d == "FH":
        annual = annual_fh_i
    elif d == "FC":
        annual = annual_fc_i
    elif d.startswith("CALENDAR"):
        annual = 12.0
    else:
        return False
    if annual <= 0:
        return False
    return (lim - slsv) <= yrs * annual


def _project_unit(unit_id, unit_type, kind, tail, position, events, util,
                  init_status, llp_rows, links, llp_rule, dates,
                  out_events, out_status, series_out, redeliv_date=None):
    n = len(dates)
    # month 0 is the as-of baseline (snapshot already reflects it): no increment.
    fh = (util.fh.copy() if util else np.zeros(n))
    fc = (util.fc.copy() if util else np.zeros(n))
    if n:
        fh[0] = 0.0
        fc[0] = 0.0
    have_util = util is not None

    # forward 12-month utilisation at each step (for "years remaining" soft tests)
    annual_fh = np.array([fh[i:i + 12].sum() for i in range(n)]) if n else np.zeros(0)
    annual_fc = np.array([fc[i:i + 12].sum() for i in range(n)]) if n else np.zeros(0)

    # Build the live event list: template events + per-engine LLPs
    live: dict[str, dict] = {}
    for e in events:
        st = init_status.get((str(unit_id), e["event_id"]), {})
        driver = e["driver"]
        start = st.get("months", 0.0) if driver.startswith("CALENDAR") else (
            st.get("fh", 0.0) if driver == "FH" else st.get("fc", 0.0))
        ev = {**e, "slsv": start, "is_llp": False}
        rm = st.get("redeliv_min")
        if rm is not None and not np.isnan(rm):     # per-unit override of the event default
            ev["redeliv_min"] = rm
        live[e["event_id"]] = ev

    llp_ids: list[str] = []
    for llp in llp_rows:
        eid = f"LLP_{llp['sn']}"
        llp_ids.append(eid)
        label = llp.get("desc") or llp["pn"]       # prefer the part description
        live[eid] = {
            "event_id": eid, "event_name": f"LLP {label} ({llp['sn']})",
            "driver": "FC", "hard_limit": llp["limit"],
            "soft_years": np.nan, "slsv": llp["slsv"], "is_llp": True,
        }

    # Build link maps from Event_Links (supports many links per source).
    hard_children: dict[str, list[str]] = {}
    soft_children: dict[str, list[str]] = {}
    for src, tgt, lt in links:
        if src in live and tgt in live:
            (hard_children if lt == "HARD" else soft_children).setdefault(src, []).append(tgt)

    # LLP bundling rule: when an engine SV happens, opportunistically replace LLPs
    # that are within their soft threshold. With NO threshold set, LLPs are NOT
    # auto-bundled — they're only replaced when they reach their own cycle limit.
    if llp_rule and llp_ids:
        link_ev, thr = llp_rule
        if link_ev in live and not np.isnan(thr):
            soft_children.setdefault(link_ev, []).extend(llp_ids)
            for lid in llp_ids:              # bundle LLPs within `thr` years of their limit
                live[lid]["soft_years"] = thr

    series = {eid: np.zeros(n) for eid in live}   # monthly SLSV snapshot per event

    # Per-component horizon: project only up to the component's redelivery date
    # (lease hand-back), not the global horizon. last_i = last in-service month.
    last_i = n - 1
    if redeliv_date is not None and n:
        rd = pd.Timestamp(redeliv_date)
        last_i = int(np.searchsorted([d.value for d in dates], rd.value, side="right")) - 1
        last_i = min(last_i, n - 1)

    for i in range(max(last_i, -1) + 1):
        date = dates[i]
        reset_now: set[str] = set()
        for eid, e in live.items():
            d = e["driver"]
            if d == "FH":
                e["slsv"] += fh[i]
            elif d == "FC":
                e["slsv"] += fc[i]
            elif d.startswith("CALENDAR"):
                e["slsv"] += 1.0 if i > 0 else 0.0
            else:
                continue
            lim = e["hard_limit"]
            # FH/FC events need utilisation to advance; skip if none
            if (d in ("FH", "FC")) and not have_util:
                continue
            if not np.isnan(lim) and lim > 0 and e["slsv"] >= lim:
                reset_now.add(eid)

        if reset_now:
            _apply_resets(live, reset_now, hard_children, soft_children,
                          annual_fh[i], annual_fc[i], date, dates,
                          unit_id, unit_type, kind, tail, position, out_events)

        # Redelivery condition: at the hand-back month, any event that would be
        # returned with less than its required min-life gets a shop visit so it
        # meets the lease return condition. These cascade HARD links only — never
        # soft links (we don't pull extra work forward just to redeliver).
        if redeliv_date is not None and i == last_i:
            redeliv_now: set[str] = set()
            for eid, e in live.items():
                if e["is_llp"]:
                    continue
                rm, lim = e.get("redeliv_min"), e["hard_limit"]
                if rm is None or np.isnan(rm) or rm <= 0 or np.isnan(lim) or lim <= 0:
                    continue
                remaining = lim - e["slsv"]
                d = e["driver"]
                required = (rm * annual_fh[i] if d == "FH"
                            else rm * annual_fc[i] if d == "FC" else rm * 12.0)
                if remaining < required:
                    redeliv_now.add(eid)
            if redeliv_now:
                _apply_resets(live, redeliv_now, hard_children, soft_children,
                              annual_fh[i], annual_fc[i], date, dates,
                              unit_id, unit_type, kind, tail, position, out_events,
                              allow_soft=False, anchor_kind="REDELIVERY")

        # record the end-of-month snapshot for every event
        for eid, e in live.items():
            series[eid][i] = e["slsv"]

    # after redelivery the component has left the fleet — freeze its last state
    for i in range(max(last_i, -1) + 1, n):
        for eid, e in live.items():
            series[eid][i] = e["slsv"]

    for eid, e in live.items():
        out_status.append({
            "Unit_ID": unit_id, "Unit_Kind": kind, "Unit_Type": unit_type,
            "Tail": tail or "SPARE", "Position": position,
            "Event": e["event_name"], "Final_SLSV": round(e["slsv"], 1),
            "Hard_Limit": e["hard_limit"],
        })
        series_out.append({
            "unit_id": unit_id, "tail": tail or "SPARE", "position": position,
            "unit_type": unit_type, "kind": kind, "event_id": e["event_id"],
            "event_name": e["event_name"], "driver": e["driver"],
            "is_llp": e["is_llp"], "series": series[eid],
        })


def _apply_resets(live, reset_now, hard_children, soft_children,
                  annual_fh_i, annual_fc_i, date, dates,
                  unit_id, unit_type, kind, tail, position, out_events,
                  allow_soft=True, anchor_kind="LIMIT"):
    # cascade HARD children (always) and SOFT children (if within threshold, and
    # only when allowed — redelivery-condition SVs never bundle soft links).
    # cascaded: child -> (trigger, source_event_id)
    cascaded: dict[str, tuple[str, str]] = {}
    frontier = list(reset_now)
    while frontier:
        p = frontier.pop()
        for child in hard_children.get(p, []):
            if child not in reset_now and child not in cascaded:
                cascaded[child] = ("HARD_LINK", p)
                frontier.append(child)
        if allow_soft:
            for child in soft_children.get(p, []):
                if (child not in reset_now and child not in cascaded
                        and _within_soft(live[child], annual_fh_i, annual_fc_i)):
                    cascaded[child] = ("SOFT_BUNDLE", p)
                    frontier.append(child)

    for eid in list(reset_now) + list(cascaded):
        e = live[eid]
        if eid in reset_now:
            if anchor_kind == "REDELIVERY":
                trigger = "REDELIVERY"
                reason = "Redelivery condition — insufficient life remaining at hand-back"
            elif e["is_llp"]:
                trigger, reason = "LLP_REPLACEMENT", "Reached its cycle limit"
            else:
                trigger, reason = "LIMIT", "Reached its limit"
        else:
            trigger, src = cascaded[eid]
            src_name = live[src]["event_name"] if src in live else src
            if trigger == "SOFT_BUNDLE":
                yrs = e.get("soft_years")
                if yrs is not None and not np.isnan(yrs):
                    months = int(round(yrs * 12))
                    reason = f"Performed earlier with {src_name} — soft link (within {months} months of its limit)"
                else:
                    reason = f"Performed earlier with {src_name} — soft link"
            else:
                reason = f"Performed earlier with {src_name} — hard link"
        out_events.append({
            "Date": date, "Tail": tail or "SPARE", "Position": position,
            "Unit_ID": unit_id, "Unit_Kind": kind, "Unit_Type": unit_type,
            "Event_ID": e["event_id"], "Event": e["event_name"],
            "Driver": e["driver"], "SLSV_at_SV": round(e["slsv"], 1),
            "Trigger": trigger, "Reason": reason,
        })
        # reset (carry remainder for usage-based; zero for calendar/link)
        lim = e["hard_limit"]
        if eid in reset_now and e["driver"] in ("FH", "FC") and not np.isnan(lim):
            e["slsv"] = max(0.0, e["slsv"] - lim)
        else:
            e["slsv"] = 0.0


# ─────────────────────────────────────────────────────────────
# Orchestration
# ─────────────────────────────────────────────────────────────

def run(ds: Dataset) -> Projection:
    as_of = _as_of(ds)
    fleet = ds.get("Fleet")
    # the Fleet sheet may also hold leased spare-engine / APU rows (rent tracking);
    # only actual aircraft rows are projected as airframes.
    if fleet is not None and not fleet.empty and "Asset_Kind" in fleet.columns:
        fleet = fleet[fleet["Asset_Kind"].astype(str).isin(["Aircraft", "", "nan", "None"])
                      | fleet["Asset_Kind"].isna()]
    units = ds.get("Components")

    # per-component redelivery date (Extension_Date if the lease is extended, else
    # Redelivery_Date). Installed engines/APU inherit their aircraft's date; spare
    # Fleet rows are keyed by serial. Each unit projects only up to this date.
    redeliv_by_unit: dict[str, pd.Timestamp] = {}
    _fleet_all = ds.get("Fleet")
    if _fleet_all is not None and not _fleet_all.empty:
        for _, _r in _fleet_all.iterrows():
            _rd = pd.to_datetime(_r.get("Extension_Date"), errors="coerce")
            if pd.isna(_rd):
                _rd = pd.to_datetime(_r.get("Redelivery_Date"), errors="coerce")
            if pd.isna(_rd):
                continue
            _key = str(_r.get("Tail") or "").strip()
            if str(_r.get("Asset_Kind") or "Aircraft") == "Aircraft":
                if _key:
                    redeliv_by_unit[_key] = _rd
                for _col in ("LHENG_ESN", "RHENG_ESN", "APU_SN"):
                    _v = _r.get(_col)
                    if pd.notna(_v) and str(_v).strip():
                        redeliv_by_unit[str(_v).strip()] = _rd
            elif _key:
                redeliv_by_unit[_key] = _rd

    # horizon end = latest in-service date across the fleet. Period 1 runs to
    # Redelivery; if a lease is extended (Extension_Date), Period 2 runs to that.
    end = as_of + relativedelta(months=DEFAULT_HORIZON_MONTHS)
    if not fleet.empty:
        _ends = []
        for _c in ("Redelivery_Date", "Extension_Date"):
            if _c in fleet.columns:
                _ends.append(pd.to_datetime(fleet[_c], errors="coerce"))
        if _ends:
            _all = pd.concat(_ends).dropna()
            if not _all.empty:
                end = max(end, _all.max() + pd.offsets.MonthEnd(0))

    dates = list(pd.date_range(as_of, end, freq="ME"))

    util_map = _build_utilisation(ds, dates, as_of)
    events_by_type = _events_by_type(ds)
    links_by_type = _links_by_type(ds)
    llp_rules = _llp_rules(ds)
    init_status = _initial_status(ds)

    # LLPs grouped by Unit_ID
    llp = ds.get("LLP_Registry")
    llp_by_unit: dict[str, list[dict]] = {}
    if not llp.empty and "Unit_ID" in llp.columns:
        for _, r in llp.iterrows():
            uid = str(r.get("Unit_ID"))
            lim = _num(r.get("Limit_Cycles"))
            slsv = _num(r.get("SLSV"), 0.0)
            if np.isnan(lim):
                continue
            desc = r.get("Description") if "Description" in llp.columns else None
            llp_by_unit.setdefault(uid, []).append({
                "pn": r.get("LLP_PN"), "sn": r.get("LLP_SN"),
                "desc": desc if (isinstance(desc, str) and desc.strip()) else None,
                "limit": lim, "slsv": 0.0 if np.isnan(slsv) else slsv,
            })

    # unit_type + kind lookups
    unit_type: dict[str, str] = {}
    unit_kind: dict[str, str] = {}
    def _cid(r):
        v = r.get("Component_ID")
        if pd.isna(v) or not str(v).strip():
            v = r.get("Unit_ID")
        return str(v)

    if not units.empty:
        for _, r in units.iterrows():
            unit_type[_cid(r)] = str(r.get("Unit_Type") or "")
            unit_kind[_cid(r)] = str(r.get("Unit_Kind") or "")

    out_events: list[dict] = []
    out_status: list[dict] = []
    out_series: list[dict] = []

    # 1) aircraft (airframe + their installed engines/APU)
    installed: set[str] = set()
    if not fleet.empty:
        for _, r in fleet.iterrows():
            tail = str(r.get("Tail"))
            fam = str(r.get("Aircraft_Family") or "")
            # airframe unit (keyed by tail)
            _project_unit(tail, fam, "AIRFRAME", tail, "AIRFRAME",
                          events_by_type.get((fam, "AIRFRAME"), []), util_map.get(tail),
                          init_status, [], links_by_type.get(fam, []), None,
                          dates, out_events, out_status, out_series,
                          redeliv_date=redeliv_by_unit.get(tail))
            # engines + APU
            for col, pos in (("LHENG_ESN", "LHENG"), ("RHENG_ESN", "RHENG"), ("APU_SN", "APU")):
                uid = r.get(col)
                if pd.isna(uid) or not str(uid).strip():
                    continue
                uid = str(uid).strip()
                installed.add(uid)
                ut = unit_type.get(uid, "")
                kind = unit_kind.get(uid) or ("APU" if pos == "APU" else "ENGINE")
                _project_unit(uid, ut, kind, tail, pos,
                              events_by_type.get((ut, kind), []), util_map.get(uid),
                              init_status, llp_by_unit.get(uid, []),
                              links_by_type.get(ut, []), llp_rules.get(ut),
                              dates, out_events, out_status, out_series,
                              redeliv_date=redeliv_by_unit.get(uid, redeliv_by_unit.get(tail)))

    # 1b) installed propellers (fly with the aircraft; LH/RH)
    if not units.empty:
        for _, r in units.iterrows():
            uid = _cid(r)
            if uid in installed or str(r.get("Unit_Kind") or "") != "PROPELLER":
                continue
            tail = str(r.get("Installed_On_Tail") or "").strip()
            if not tail:
                continue
            installed.add(uid)
            ut = str(r.get("Unit_Type") or "")
            _project_unit(uid, ut, "PROPELLER", tail, str(r.get("Position") or ""),
                          events_by_type.get((ut, "PROPELLER"), []), util_map.get(uid),
                          init_status, [], links_by_type.get(ut, []), None,
                          dates, out_events, out_status, out_series,
                          redeliv_date=redeliv_by_unit.get(uid, redeliv_by_unit.get(tail)))

    # 1c) installed landing gear (belongs to the tail; nose + MLG LH/RH)
    if not units.empty:
        for _, r in units.iterrows():
            uid = _cid(r)
            if uid in installed or str(r.get("Unit_Kind") or "") != "LANDING_GEAR":
                continue
            tail = str(r.get("Installed_On_Tail") or "").strip()
            if not tail:
                continue
            installed.add(uid)
            ut = str(r.get("Unit_Type") or "")
            _project_unit(uid, ut, "LANDING_GEAR", tail, str(r.get("Position") or "LG"),
                          events_by_type.get((ut, "LANDING_GEAR"), []), util_map.get(tail),
                          init_status, [], links_by_type.get(ut, []), None,
                          dates, out_events, out_status, out_series,
                          redeliv_date=redeliv_by_unit.get(uid, redeliv_by_unit.get(tail)))

    # 2) spares (units in the Units sheet not installed on any aircraft)
    if not units.empty:
        for _, r in units.iterrows():
            uid = _cid(r)
            if uid in installed:
                continue
            pos = str(r.get("Position") or "SPARE")
            if pos != "SPARE":
                continue
            ut = str(r.get("Unit_Type") or "")
            kind = str(r.get("Unit_Kind") or "")
            _project_unit(uid, ut, kind, None, "SPARE",
                          events_by_type.get((ut, kind), []), util_map.get(uid),
                          init_status, llp_by_unit.get(uid, []),
                          links_by_type.get(ut, []), llp_rules.get(ut),
                          dates, out_events, out_status, out_series,
                          redeliv_date=redeliv_by_unit.get(uid))

    events_df = pd.DataFrame(out_events)
    if not events_df.empty:
        events_df = events_df.sort_values("Date").reset_index(drop=True)
    status_df = pd.DataFrame(out_status)

    v1, v2 = _build_sv_plans(events_df, dates, fleet)
    frames = _build_aircraft_frames(out_series, util_map, dates, fleet)

    return Projection(events=events_df, unit_status=status_df, dates=dates,
                      horizon_end=end, sv_plan_v1=v1, sv_plan_v2=v2,
                      aircraft_frames=frames)


# ─────────────────────────────────────────────────────────────
# Structured outputs: SV calendar (v1), event×aircraft matrix (v2),
# per-aircraft monthly counter sheets.
# ─────────────────────────────────────────────────────────────

def _build_sv_plans(events_df, dates, fleet):
    tails = list(fleet["Tail"].dropna().astype(str)) if not fleet.empty else []
    v1 = pd.DataFrame({"DATE": dates})
    for t in tails:
        v1[t] = None
    if events_df is None or events_df.empty:
        return v1, pd.DataFrame({"Shop Visits": []})

    ev = events_df.copy()
    ev = ev[ev["Tail"].isin(tails)]
    ev["Month"] = pd.to_datetime(ev["Date"]) + pd.offsets.MonthEnd(0)
    ev["Label"] = [_sv_label(p, e, t == "LLP_REPLACEMENT" or str(ename).startswith("LLP"), uid)
                   for p, e, t, ename, uid in
                   zip(ev["Position"], ev["Event"], ev["Trigger"], ev["Event"], ev["Unit_ID"])]

    # v1: month × aircraft calendar (dedup labels per cell)
    date_to_row = {d: i for i, d in enumerate(dates)}
    for (month, tail), grp in ev.groupby(["Month", "Tail"]):
        ri = date_to_row.get(month)
        if ri is None or tail not in v1.columns:
            continue
        labels = sorted(set(grp["Label"]))
        v1.at[ri, tail] = ", ".join(labels)

    # v2: event label × aircraft, cell = sorted unique dates
    labels_all = sorted(ev["Label"].unique())
    v2 = pd.DataFrame({"Shop Visits": labels_all})
    idx = {lab: i for i, lab in enumerate(labels_all)}
    for t in tails:
        v2[t] = None
    for (label, tail), grp in ev.groupby(["Label", "Tail"]):
        if tail not in v2.columns:
            continue
        ds = sorted({pd.to_datetime(d).strftime("%Y-%m-%d") for d in grp["Date"]})
        v2.at[idx[label], tail] = ", ".join(ds)
    return v1, v2


def _build_aircraft_frames(out_series, util_map, dates, fleet):
    """One monthly DataFrame per aircraft: utilisation + every counter over time."""
    frames: dict[str, pd.DataFrame] = {}
    by_tail: dict[str, list[dict]] = {}
    for s in out_series:
        by_tail.setdefault(s["tail"], []).append(s)

    for tail, series_list in by_tail.items():
        if tail == "SPARE":
            continue
        data = {"DATE": dates}
        # utilisation columns (airframe + each installed unit)
        seen_util = set()
        for s in series_list:
            uid, pos = s["unit_id"], s["position"]
            if uid in seen_util:
                continue
            seen_util.add(uid)
            u = util_map.get(uid)
            if u is not None:
                tag = "REG" if pos == "AIRFRAME" else pos
                data[f"{tag}-UTIL FH"] = np.round(u.fh, 1)
                data[f"{tag}-UTIL FC"] = np.round(u.fc, 1)
        # counter columns
        for s in series_list:
            pos = s["position"]
            col = s["event_name"] if pos == "AIRFRAME" else f"{pos} {s['event_name']}"
            data[f"{col} SLSV"] = np.round(s["series"], 1)
        frames[tail] = pd.DataFrame(data)
    return frames
