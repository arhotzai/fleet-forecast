"""
Fleet ERP — Build the relational database workbook FROM the legacy DATA INPUT.xlsx
==================================================================================
Reads the user's existing wide-format workbook and emits the new, fully-relational
FLEET_ERP_DATABASE.xlsx (unit-based: aircraft = airframe + serialized engines/APU,
each unit with its own utilisation; LLPs sourced from the company DB export).

This is a one-time migration, so it DOES parse the legacy column names here — the
whole point is that the new file never needs to.

Run:   python3 tools/build_database.py "/Users/telischotzai/Desktop/DATA INPUT.xlsx"
Out:   data/FLEET_ERP_DATABASE.xlsx
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

# ─────────────────────────────────────────────────────────────
# Controlled vocabularies (dropdown sources)
# ─────────────────────────────────────────────────────────────
ENUMS: dict[str, list[str]] = {
    "UnitKind":              ["AIRFRAME", "ENGINE", "APU", "PROPELLER"],
    "Driver":                ["FH", "FC", "CALENDAR_MONTHS", "CALENDAR_DAYS"],
    "SoftThresholdType":     ["PCT", "ABS_FH", "ABS_FC", "ABS_MONTHS", "NONE"],
    "RedeliveryMinLifeType": ["FH", "FC", "MONTHS", "NONE"],
    "LinkType":              ["HARD", "SOFT_BUNDLE", "NONE"],
    "WorkscopeClass":        ["LIGHT", "HEAVY", "LLP_REPLACEMENT", "NONE"],
    "AircraftStatus":        ["ACTIVE", "AOG", "STORED"],
    "LeasedOwned":           ["Leased", "Owned"],
    "Position":              ["LHENG", "RHENG", "APU", "LHPROP", "RHPROP", "SPARE"],
    "ProjectionEndMode":     ["REDELIVERY", "FIXED_DATE", "HORIZON_MONTHS"],
    "ScopeType":             ["ALL", "CATEGORY", "COMPONENT"],
    "MethodType":            ["MR_STANDARD", "MR_FH_FC_RATIO", "PBH_FLAT", "TM_WORKSCOPE", "EOL_PRORATA"],
    "RateBasis":             ["PER_FH", "PER_FC", "PER_MONTH"],
    "CompBasis":             ["PER_REMAINING_FH", "PER_CONSUMED_FH", "PER_REMAINING_FC", "PER_CONSUMED_FC"],
    "EventType":             ["LIGHT_SV", "HEAVY_SV", "LLP_REPLACEMENT"],
    "Bool":                  ["TRUE", "FALSE"],
}
_LIST_COL = {k: i + 1 for i, k in enumerate(ENUMS.keys())}

# Which legacy sheets are engine / APU families (vs airframe)
ENGINE_FAMILIES = {"V2500", "PW1100", "PW127"}
APU_FAMILIES = {"APU"}

# ─────────────────────────────────────────────────────────────
# Styling
# ─────────────────────────────────────────────────────────────
HEADER_FILL = PatternFill("solid", fgColor="1F3864")
HEADER_FONT = Font(color="FFFFFF", bold=True, size=11)
TITLE_FONT = Font(bold=True, size=14, color="1F3864")
NOTE_FONT = Font(size=10, color="404040")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _style_header(ws, ncols):
    for ci in range(1, ncols + 1):
        c = ws.cell(row=1, column=ci)
        c.fill, c.font, c.border = HEADER_FILL, HEADER_FONT, BORDER
        c.alignment = Alignment(vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 28
    ws.freeze_panes = "A2"


def _autoscale(ws, headers, min_w=12, max_w=40):
    for ci, h in enumerate(headers, start=1):
        ws.column_dimensions[get_column_letter(ci)].width = max(min_w, min(max_w, len(str(h)) + 4))


def _add_dropdown(ws, col_idx, enum_key, first=2, last=2000):
    vals = ENUMS[enum_key]
    cl = get_column_letter(_LIST_COL[enum_key])
    rng = f"_Lists!${cl}$2:${cl}${1 + len(vals)}"
    dv = DataValidation(type="list", formula1=rng, allow_blank=True)
    dv.errorTitle, dv.error = "Invalid value", f"Allowed: {', '.join(vals)}"
    ws.add_data_validation(dv)
    L = get_column_letter(col_idx)
    dv.add(f"{L}{first}:{L}{last}")


def _sheet(wb, name, headers, rows):
    ws = wb.create_sheet(name)
    ws.append(headers)
    for r in rows:
        ws.append(r)
    _style_header(ws, len(headers))
    _autoscale(ws, headers)
    return ws


# ─────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────

def _clean(v):
    if pd.isna(v):
        return None
    if isinstance(v, str):
        return v.strip()
    return v


def _d(v):
    """Date → YYYY-MM-DD string."""
    if pd.isna(v):
        return None
    try:
        return pd.to_datetime(v).strftime("%Y-%m-%d")
    except Exception:
        return _clean(v)


def _driver_from_legacy(drv: str | None, event_name: str) -> str:
    """Normalise a legacy DRIVERS string / event name to a clean Driver enum."""
    s = str(drv or "").upper()
    if "FH" in s:
        return "FH"
    if "FC" in s:
        return "FC"
    if "CALENDAR" in s or "MONTH" in s:
        return "CALENDAR_MONTHS"
    # Infer from the event name when the template left DRIVERS blank
    n = event_name.upper()
    if n.endswith("FH SLSV") or "-FH" in n:
        return "FH"
    if n.endswith("FC SLSV") or "-FC" in n:
        return "FC"
    if re.search(r"\d+Y\b", n) or "MT" in n:        # 6Y, 12Y, 24MT …  → calendar
        return "CALENDAR_MONTHS"
    if re.search(r"\d+K\b", n) or n.endswith("C"):  # 18K, 36K, 1C, 2C → cycles
        return "FC"
    return "FH"


def _strip_pos(name: str) -> str:
    """Make an engine/apu event position-agnostic: 'ENG1-FH SLSV' -> 'FH'."""
    n = re.sub(r"^(ENG1|ENG2|APU)[- ]?", "", name).strip()
    n = re.sub(r"\s*SLSV$", "", n).strip()
    n = re.sub(r"\s*-\s*$", "", n).strip()
    return n or name


def _drop_driver_suffix(name: str) -> str:
    """Drop a trailing FH/FC driver token from a sub-event name — the driver is a
    separate field now. 'HSI-FH' -> 'HSI', 'RGB FH' -> 'RGB', but bare 'FH'/'FC'
    main counters are left intact (no leading separator to match)."""
    return re.sub(r"[-\s]+F[HC]$", "", str(name)).strip()


def _clean_name(name: str) -> str:
    """Drop a trailing 'SLSV' and any dangling separators: '6Y-SLSV' -> '6Y'."""
    n = re.sub(r"\s*SLSV$", "", str(name))
    return re.sub(r"[\s\-]+$", "", n).strip()


def _strip_event_id(name: str) -> str:
    return _clean_name(name).replace("-", "_").replace(" ", "_").strip("_")


# ─────────────────────────────────────────────────────────────
# Parse the legacy template sheets into clean event definitions
# ─────────────────────────────────────────────────────────────

def parse_templates(xl: pd.ExcelFile):
    """
    Returns three dicts keyed by Unit_Type:
      airframe_events[type] / engine_events[type] / apu_events[type]
    each a list of dict(event_id, event_name, driver, hard_limit, link_to, soft_thr).
    """
    airframe, engine, apu = {}, {}, {}
    tmpl_sheets = [s for s in xl.sheet_names if s.endswith("_columns")]

    for sheet in tmpl_sheets:
        fam = sheet[:-len("_columns")]
        df = pd.read_excel(xl, sheet)
        famcol = df.columns[0]
        is_eng = fam in ENGINE_FAMILIES
        is_apu = fam in APU_FAMILIES

        # Build a limit lookup from the *-LIMIT rows (position-stripped for eng/apu)
        limit_lookup: dict[str, float] = {}
        for _, r in df.iterrows():
            nm = _clean(r[famcol])
            lim = r.get("LIMIT")
            if isinstance(nm, str) and nm.upper().endswith("LIMIT") and pd.notna(lim):
                key = re.sub(r"\s*LIMIT$", "", nm)
                if is_eng or is_apu:
                    key = re.sub(r"^(ENG1|ENG2|APU)[- ]?", "", key)
                _k = _clean_name(key)
                if is_eng or is_apu:
                    _k = _drop_driver_suffix(_k)
                limit_lookup[_k.upper()] = float(lim)

        seen = set()
        for _, r in df.iterrows():
            nm = _clean(r[famcol])
            if not isinstance(nm, str) or "SLSV" not in nm.upper():
                continue
            if "LLP" in nm.upper():
                continue  # LLPs come from the company DB, not the template
            # On an AIRFRAME template, skip the engine/APU rows — those belong to
            # the engine/APU unit types, not the airframe.
            if not (is_eng or is_apu) and re.match(r"^(ENG1|ENG2|APU)\b", nm.upper()):
                continue
            if is_eng or is_apu:
                ev_name = _drop_driver_suffix(_strip_pos(nm))  # collapse pos; drop FH/FC token
                base_key = _drop_driver_suffix(
                    _clean_name(re.sub(r"^(ENG1|ENG2|APU)[- ]?", "", nm))).upper()
            else:
                ev_name = _clean_name(nm)
                base_key = ev_name.upper()
            if ev_name in seen:
                continue
            seen.add(ev_name)

            hard_limit = limit_lookup.get(base_key) or limit_lookup.get((ev_name + " ").strip().upper())
            # Legacy templates allow up to three links per event (SV1/SV2/SV3).
            links = []
            for k in ("LINKED SV1", "LINKED SV2", "LINKED SV3"):
                lv = _clean(r.get(k))
                if isinstance(lv, str) and lv.strip():
                    tgt = _strip_event_id(_drop_driver_suffix(_strip_pos(lv)))
                    if tgt:
                        links.append(tgt)
            thr = r.get("THRESHOLD")
            soft_thr = float(thr) if pd.notna(thr) else None

            rec = {
                "event_id": _strip_event_id(ev_name),
                "event_name": ev_name,
                "driver": _driver_from_legacy(r.get("DRIVERS"), nm),
                "hard_limit": hard_limit,
                "links": links,
                "link_to": links[0] if links else None,
                "soft_thr": soft_thr,
            }
            tgt = engine if is_eng else apu if is_apu else airframe
            tgt.setdefault(fam, []).append(rec)

    return airframe, engine, apu


# ─────────────────────────────────────────────────────────────
# Build
# ─────────────────────────────────────────────────────────────

def build(src: Path, llp_report: Path | None = None) -> Path:
    xl = pd.ExcelFile(src)
    ac = pd.read_excel(xl, "AC_list")
    util = pd.read_excel(xl, "UTIL")
    cl = pd.read_excel(xl, "COMPONENT_LIMITS")

    # Live LLP status export from the company database (authoritative for LLPs).
    rep = None
    if llp_report and Path(llp_report).exists():
        rep = pd.read_excel(llp_report, header=0)
        rep.columns = ["eng_partno", "eng_esn", "llp_partno", "llp_sn",
                       "life_limit", "togo", "counter", "expected_date"]

    airframe_ev, engine_ev, apu_ev = parse_templates(xl)

    # Propellers: ACTUATOR / NUT / TRANSFER TUBE are PROPELLER rotables (ATRs), not
    # airframe events. Pull them out of the airframe events into a PROPELLER unit
    # type (FH-driven, LH/RH), collapsing the LH/RH (1/2) and the -FH suffix.
    _PROP_RE = re.compile(r"\b(ACTUATOR|NUT|TRANSFER\s*TUBE|HUB)\b", re.I)
    _prop_ev: dict[str, dict] = {}
    for fam in list(airframe_ev):
        kept = []
        for e in airframe_ev[fam]:
            if _PROP_RE.search(e["event_name"]):
                base = re.sub(r"\s*\d+\s*", " ", e["event_name"])     # drop LH/RH digit
                base = _drop_driver_suffix(base).strip(" -")
                eid = _strip_event_id(base)
                if eid and eid not in _prop_ev:
                    _prop_ev[eid] = {"event_id": eid, "event_name": base, "driver": "FH",
                                     "hard_limit": e.get("hard_limit"), "links": [],
                                     "link_to": None, "soft_thr": None}
            else:
                kept.append(e)
        airframe_ev[fam] = kept
    propeller_ev = {"PROP": list(_prop_ev.values())} if _prop_ev else {}

    wb = Workbook()
    lists = wb.active
    lists.title = "_Lists"
    for key, col in _LIST_COL.items():
        lists.cell(row=1, column=col, value=key).font = Font(bold=True)
        for i, v in enumerate(ENUMS[key], start=2):
            lists.cell(row=i, column=col, value=v)
    lists.sheet_state = "hidden"

    _build_instructions(wb.create_sheet("Instructions"))

    # ---- Config ----
    _sheet(wb, "Config", ["Setting", "Value", "Notes"], [
        ["As_Of_Date", _d(ac["DATE"].dropna().iloc[0]) if "DATE" in ac else "2026-07-20", "Snapshot date the projection starts from."],
        ["Projection_End_Mode", "REDELIVERY", "REDELIVERY | FIXED_DATE | HORIZON_MONTHS"],
        ["Projection_End_Value", "", "Date or month-count; blank for REDELIVERY."],
        ["Base_Currency", "USD", "Reporting currency."],
        ["Schema_Version", "2.0", "Do not edit."],
    ])

    # ---- Unit_Types (per-type LLP behaviour) ----
    # Include every airframe family seen in the fleet, even those without an
    # event template yet (e.g. ATR42) so the user can add events later.
    airframe_families_all = {
        _clean(r.get("FAMILY")) for _, r in ac.iterrows()
        if _clean(r.get("FAMILY")) and _clean(r.get("FAMILY")) not in ENGINE_FAMILIES
        and _clean(r.get("FAMILY")) not in APU_FAMILIES
    }
    ut_rows = []
    for fam in sorted(airframe_families_all):
        note = "Airframe type — no LLPs." + ("" if fam in airframe_ev else "  (no event template in source — add events.)")
        ut_rows.append(["AIRFRAME", fam, "NONE", "", "", note])
    for fam in sorted(engine_ev):
        # LLP link target must be an actual Event_ID of this engine type.
        ev_ids = {e["event_id"] for e in engine_ev[fam]}
        link = "HSI" if "HSI" in ev_ids else ("FH" if "FH" in ev_ids else "")
        ut_rows.append(["ENGINE", fam, "FC", link, 1.0, "Engine type — LLPs driven by FC, bundled with the primary engine SV."])
    for fam in sorted(apu_ev):
        ut_rows.append(["APU", fam, "FC", "FH", 1.0, "APU type — LLPs driven by FC."])
    for fam in sorted(propeller_ev):
        ut_rows.append(["PROPELLER", fam, "NONE", "", "",
                        "Propeller rotables (actuator / nut / transfer tube) — FH-driven, LH/RH."])
    ut = _sheet(wb, "Unit_Types",
                ["Unit_Kind", "Unit_Type", "LLP_Driver", "LLP_Link_To_Event",
                 "LLP_Soft_Threshold_Years", "Notes"],
                ut_rows)
    _add_dropdown(ut, 1, "UnitKind")
    _add_dropdown(ut, 3, "Driver", first=2, last=200)

    # ---- Events (ONE relational sheet for all kinds; Unit_Kind tags each row) ----
    # Soft-bundling thresholds are left blank: the legacy THRESHOLD used a different
    # semantic we're replacing, so the user sets these fresh. Links live in Event_Links.
    # Soft_Threshold_Years: bundle this event if remaining life < years × annual utilisation.
    ev_headers = ["Unit_Kind", "Unit_Type", "Event_ID", "Event_Name", "Driver", "Hard_Limit",
                  "Soft_Threshold_Years", "Redelivery_Min_Life_Years", "Workscope_Class"]
    ev_rows = []
    for kind, evmap in (("AIRFRAME", airframe_ev), ("ENGINE", engine_ev),
                        ("APU", apu_ev), ("PROPELLER", propeller_ev)):
        for fam in sorted(evmap):
            for e in evmap[fam]:
                ev_rows.append([kind, fam, e["event_id"], e["event_name"], e["driver"],
                                e["hard_limit"], "", "", "HEAVY"])
    ws = _sheet(wb, "Events", ev_headers, ev_rows)
    _add_dropdown(ws, 1, "UnitKind")
    _add_dropdown(ws, 5, "Driver")
    _add_dropdown(ws, 9, "WorkscopeClass")

    # ---- Event_Links (many-to-many: one row per link; an event can link to several) ----
    el_rows = []
    for evmap in (airframe_ev, engine_ev, apu_ev):
        for fam in sorted(evmap):
            for e in evmap[fam]:
                for tgt in e.get("links", []):
                    el_rows.append([fam, e["event_id"], tgt, "HARD"])
    ws = _sheet(wb, "Event_Links",
                ["Unit_Type", "Source_Event", "Target_Event", "Link_Type"], el_rows)
    _add_dropdown(ws, 4, "LinkType")

    # ---- Fleet (airframe units = aircraft) ----
    fleet_headers = ["Tail", "Aircraft_Family", "MSN", "Operator", "Owner", "Leased_Owned",
                     "Status", "Delivery_Date", "Redelivery_Date", "Extension_Date",
                     "LHENG_ESN", "RHENG_ESN", "APU_SN", "Financial_Scenario_ID"]
    fleet_rows, fleet_tails = [], []
    for _, r in ac.iterrows():
        fam = _clean(r.get("FAMILY"))
        # An aircraft is any row whose family is NOT an engine/APU type.
        if fam in ENGINE_FAMILIES or fam in APU_FAMILIES or not fam:
            continue
        tail = _clean(r.get("REG"))
        if not tail:
            continue
        fleet_tails.append(tail)
        fleet_rows.append([
            tail, fam, _clean(r.get("MSN/ESN")), _clean(r.get("OPERATOR")),
            _clean(r.get("OWNER")), _clean(r.get("LEASED/OWNED")), "ACTIVE",
            _d(r.get("DELIVERY")), _d(r.get("REDELIVERY")), _d(r.get("EXTENSION")),
            _clean(r.get("ENG1")), _clean(r.get("ENG2")), _clean(r.get("APU")),
            "FIN_DEFAULT",
        ])
    fleet = _sheet(wb, "Fleet", fleet_headers, fleet_rows)
    _add_dropdown(fleet, 6, "LeasedOwned")
    _add_dropdown(fleet, 7, "AircraftStatus")

    # ---- Units (serialized engines + APUs, incl. spares) ----
    engine_type: dict[str, str] = {}   # ESN -> engine type
    apu_serials: set[str] = set()
    install: dict[str, tuple[str, str]] = {}   # serial -> (tail, position)
    for _, r in ac.iterrows():
        fam = _clean(r.get("FAMILY"))
        tail = _clean(r.get("REG"))
        # Some AC_list rows are standalone engines/APUs (REG = the serial itself).
        # Use them to recover a spare unit's TYPE, but never as an installation.
        if fam in ENGINE_FAMILIES:
            if tail:
                engine_type[str(tail)] = fam
            continue
        if fam in APU_FAMILIES:
            if tail:
                apu_serials.add(str(tail))
            continue
        for esn_c, ty_c, pos in (("ENG1", "ENG1-TYPE", "LHENG"), ("ENG2", "ENG2-TYPE", "RHENG")):
            esn = _clean(r.get(esn_c))
            if esn:
                engine_type.setdefault(str(esn), _clean(r.get(ty_c)))
                install[str(esn)] = (tail, pos)
        apu = _clean(r.get("APU"))
        if apu:
            apu_serials.add(str(apu))
            install[str(apu)] = (tail, "APU")

    # Classify any extra serials from UTIL and the LLP DB via their TYPE tag
    for _, r in util.iterrows():
        uid = _clean(r.get("REG"))
        t = str(_clean(r.get("TYPE")) or "")
        if not uid:
            continue
        if t == "ENG":
            engine_type.setdefault(str(uid), None)
        elif t == "APU":
            apu_serials.add(str(uid))
    for _, r in cl.iterrows():
        esn = _clean(r.get("REG"))
        if not esn or str(esn) in ("0", "0.0"):
            continue
        if str(_clean(r.get("TYPE"))) == "APU":
            apu_serials.add(str(esn))
        else:
            engine_type.setdefault(str(esn), None)

    # Serials that appear only in the LLP report = units we don't otherwise know
    # about (e.g. newly acquired spares). Add them so their LLPs aren't orphaned.
    known = set(engine_type) | apu_serials
    if rep is not None:
        for esn in rep["eng_esn"].dropna().astype(str).unique():
            if esn in known:
                continue
            if re.match(r"^P-\d", esn):      # P-#### pattern = APU serial
                apu_serials.add(esn)
            else:
                engine_type.setdefault(esn, None)

    unit_rows = []
    for sn in sorted(apu_serials):
        tail, pos = install.get(sn, (None, "SPARE"))
        unit_rows.append([sn, "APU", "APU", tail or "", pos if tail else "SPARE"])
    for esn in sorted(set(engine_type) - apu_serials):
        tail, pos = install.get(esn, (None, "SPARE"))
        unit_rows.append([esn, "ENGINE", engine_type.get(esn) or "", tail or "", pos if tail else "SPARE"])
    # Propeller units: LH/RH per aircraft that carries propeller data (ATRs)
    prop_tails = []
    if propeller_ev:
        for _, r in ac.iterrows():
            fam = _clean(r.get("FAMILY"))
            tail = _clean(r.get("REG"))
            if not tail or fam in ENGINE_FAMILIES or fam in APU_FAMILIES:
                continue
            has_prop = any(pd.notna(r.get(f"{x} {i}-FH SLSV"))
                           for x in ("ACTUATOR", "NUT", "TRANSFER TUBE") for i in ("1", "2"))
            if not has_prop:
                continue
            prop_tails.append(str(tail))
            for pos, lbl in (("LHPROP", "LH"), ("RHPROP", "RH")):
                unit_rows.append([f"{tail}-PROP-{lbl}", "PROPELLER", "PROP", str(tail), pos])
    units = _sheet(wb, "Components",
                   ["Component_ID", "Unit_Kind", "Unit_Type", "Installed_On_Tail", "Position"],
                   unit_rows)
    _add_dropdown(units, 2, "UnitKind")
    _add_dropdown(units, 5, "Position")

    # ---- Utilization (per unit; ANNUAL value + 12 monthly % in ONE sheet) ----
    # Annual_FH/FC is the yearly total; the month columns are percentages that
    # distribute it across the year (they sum to ~1.0). Monthly FH = Annual_FH × month%.
    MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
              "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    util_headers = ["Unit_ID", "Unit_Kind", "Annual_FH", "Annual_FC", "FH_FC_Ratio",
                    "Annual_Evolution_Pct", "Min_Annual_FH", "Max_Annual_FH",
                    "AOG_Lead_In_Months"] + MONTHS
    kind_map = {"AC": "AIRFRAME", "ENG": "ENGINE", "APU": "APU"}
    util_rows = []
    for _, r in util.iterrows():
        uid = _clean(r.get("REG"))
        if not uid:
            continue
        kind = kind_map.get(str(_clean(r.get("TYPE"))), "")
        fhfc = r.get("FH/FC")
        fh = r.get("UTIL FH")
        fc = r.get("UTIL FC")
        monthly = [round(float(r.get(m)), 3) if pd.notna(r.get(m)) else None for m in range(1, 13)]
        util_rows.append([
            uid, kind,
            round(float(fh), 1) if pd.notna(fh) else None,
            round(float(fc), 1) if pd.notna(fc) else None,
            round(float(fhfc), 3) if pd.notna(fhfc) else None,
            float(r.get("EVOLUTION")) if pd.notna(r.get("EVOLUTION")) else 0.0,
            float(r.get("MIN")) if pd.notna(r.get("MIN")) else None,
            float(r.get("MAX")) if pd.notna(r.get("MAX")) else None,
            int(r.get("AOG")) if pd.notna(r.get("AOG")) else 0,
        ] + monthly)
    # Propeller utilisation = the airframe's (propellers fly with the aircraft).
    if propeller_ev and prop_tails:
        _af_util = {str(row[0]): row for row in util_rows if row[1] == "AIRFRAME"}
        for tail in prop_tails:
            base = _af_util.get(str(tail))
            if base is None:
                continue
            for lbl in ("LH", "RH"):
                util_rows.append([f"{tail}-PROP-{lbl}", "PROPELLER"] + list(base[2:]))
    uw = _sheet(wb, "Utilization", util_headers, util_rows)
    _add_dropdown(uw, 2, "UnitKind")

    # ---- LLP_Registry (from the live LLP report; SLSV = Limit - ToGo) ----
    # Descriptions aren't in the report. A description is a property of the PART
    # NUMBER (not the serial), so join from COMPONENT_LIMITS by PN → 100% coverage.
    desc_by_pn = {}
    for _, r in cl.iterrows():
        pn = _clean(r.get("PN"))
        d = _clean(r.get("DESCRIPTION"))
        if pn and isinstance(d, str) and d.strip() and d.strip().lower() != "nan":
            desc_by_pn.setdefault(str(pn).strip(), d.strip())

    unit_kind_lookup = {u[0]: u[1] for u in unit_rows}
    llp_headers = ["Unit_ID", "Unit_Kind", "LLP_PN", "LLP_SN", "Description",
                   "Counter", "Limit_Cycles", "ToGo", "SLSV", "Expected_Date"]
    llp_rows = []
    if rep is not None:
        for _, r in rep.iterrows():
            esn = _clean(r.get("eng_esn"))
            if not esn:
                continue
            lim = r.get("life_limit")
            togo = r.get("togo")
            slsv = (float(lim) - float(togo)) if pd.notna(lim) and pd.notna(togo) else None
            sn = _clean(r.get("llp_sn"))
            pn = _clean(r.get("llp_partno"))
            llp_rows.append([
                esn, unit_kind_lookup.get(str(esn), ""),
                pn, sn,
                desc_by_pn.get(str(pn).strip()) if pn else None,
                _clean(r.get("counter")),
                float(lim) if pd.notna(lim) else None,
                float(togo) if pd.notna(togo) else None,
                round(slsv, 2) if slsv is not None else None,
                pd.to_datetime(r.get("expected_date"), dayfirst=True, errors="coerce").strftime("%Y-%m-%d")
                if pd.notna(r.get("expected_date")) else None,
            ])
    else:
        # Fallback: limits only, from COMPONENT_LIMITS
        for _, r in cl.iterrows():
            esn = _clean(r.get("REG"))
            if not esn or str(esn) in ("0", "0.0"):
                continue
            llp_rows.append([esn, "", None, _clean(r.get("SN")), _clean(r.get("DESCRIPTION")),
                             "C", float(r.get("LIMIT")) if pd.notna(r.get("LIMIT")) else None,
                             None, None, None])
    _sheet(wb, "LLP_Registry", llp_headers, llp_rows)

    # ---- Unit_Status (current counters, best-effort from wide AC_list) ----
    _build_unit_status(wb, ac, airframe_ev, engine_ev, apu_ev, fleet_tails)

    # ---- Financial scaffolding (no real contract data provided yet) ----
    _build_financial(wb)

    out = Path(__file__).resolve().parents[1] / "data" / "FLEET_MAINTENANCE_STATUS_DATABASE.xlsx"
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    # Always keep an easy-to-find copy on the Desktop for the user.
    desktop = Path.home() / "Desktop" / "FLEET_MAINTENANCE_STATUS_DATABASE.xlsx"
    try:
        import shutil
        shutil.copyfile(out, desktop)
    except Exception:
        desktop = None
    return out, {
        "desktop_copy": str(desktop) if desktop else "(could not write)",
        "aircraft": len(fleet_rows), "units": len(unit_rows),
        "airframe_events": sum(len(v) for v in airframe_ev.values()),
        "engine_events": sum(len(v) for v in engine_ev.values()),
        "apu_events": sum(len(v) for v in apu_ev.values()),
        "util_rows": len(util_rows), "llp_rows": len(llp_rows),
        "llp_source": "live report" if rep is not None else "COMPONENT_LIMITS fallback",
        "families": sorted(set(airframe_ev) | set(engine_ev) | set(apu_ev)),
    }


def _build_unit_status(wb, ac, airframe_ev, engine_ev, apu_ev, fleet_tails):
    """Map legacy wide SLSV / TSN / CSN columns into long Unit_Status rows."""
    headers = ["Unit_ID", "Event_ID", "Driver", "SLSV_FH", "SLSV_FC", "SLSV_Months",
               "TSN", "CSN", "Last_SV_Date"]
    # driver lookup for airframe events: (family, event_id) -> driver
    af_driver = {(fam, e["event_id"]): e["driver"]
                 for fam, evs in airframe_ev.items() for e in evs}
    rows = []
    cols = list(ac.columns)
    for _, r in ac.iterrows():
        tail = _clean(r.get("REG"))
        if tail not in fleet_tails:
            continue
        fam = _clean(r.get("FAMILY"))
        # Airframe checks: columns like '6Y-SLSV', '12Y-SLSV', 'LG-SLSV', '36K-SLSV' …
        for c in cols:
            if not isinstance(c, str) or not c.upper().endswith("-SLSV"):
                continue
            if c.upper().startswith(("ENG1", "ENG2", "APU")):
                continue
            val = r.get(c)
            if pd.isna(val):
                continue
            ev_id = _strip_event_id(re.sub(r"-SLSV$", "", c, flags=re.I))
            driver = af_driver.get((fam, ev_id)) or (
                "CALENDAR_MONTHS" if re.search(r"\d+Y|MT", c.upper()) else "FC")
            is_months = driver.startswith("CALENDAR")
            rows.append([tail, ev_id, driver,
                         None, None if is_months else val, val if is_months else None,
                         _clean(r.get("REG-TSN")), _clean(r.get("REG-CSN")), None])
        # Engine FH/FC SLSV mapped onto the installed engine serial (driver = FH)
        for pos in ("ENG1", "ENG2"):
            esn = _clean(r.get(pos))
            if not esn:
                continue
            _tsn, _csn = _clean(r.get(f"{pos}-TSN")), _clean(r.get(f"{pos}-CSN"))
            fh = r.get(f"{pos}-FH SLSV")
            fc = r.get(f"{pos}-FC SLSV")
            if pd.notna(fh) or pd.notna(fc):
                rows.append([esn, "FH", "FH", fh if pd.notna(fh) else None,
                             fc if pd.notna(fc) else None, None, _tsn, _csn, None])
            # engine sub-events HSI / CSI / RGB (driver FH) — kept without the -FH suffix
            for sub in ("HSI", "CSI", "RGB"):
                sv = r.get(f"{pos} {sub}-FH SLSV")
                if pd.notna(sv):
                    rows.append([esn, sub, "FH", sv, None, None, _tsn, _csn, None])
        # APU — driver is APU FH
        apu = _clean(r.get("APU"))
        if apu:
            fh, fc = r.get("APU-FH SLSV"), r.get("APU-FC SLSV")
            if pd.notna(fh) or pd.notna(fc):
                rows.append([apu, "FH", "FH", fh if pd.notna(fh) else None,
                             fc if pd.notna(fc) else None, None,
                             _clean(r.get("APU-TSN")), _clean(r.get("APU-CSN")), None])
        # Propellers — ACTUATOR / NUT / TRANSFER TUBE SLSV → LH(1)/RH(2) prop units (FH)
        _rtsn, _rcsn = _clean(r.get("REG-TSN")), _clean(r.get("REG-CSN"))
        for lbl, idx in (("LH", "1"), ("RH", "2")):
            for ev_col, ev_id in (("ACTUATOR", "ACTUATOR"), ("NUT", "NUT"),
                                  ("TRANSFER TUBE", "TRANSFER_TUBE")):
                sv = r.get(f"{ev_col} {idx}-FH SLSV")
                if pd.notna(sv):
                    rows.append([f"{tail}-PROP-{lbl}", ev_id, "FH", sv, None, None,
                                 _rtsn, _rcsn, None])
    _sheet(wb, "Unit_Status", headers, rows)


def _build_financial(wb):
    _sheet(wb, "Financial_Scenarios",
           ["Financial_Scenario_ID", "Name", "Currency", "Escalation_Base_Date", "Notes"],
           [["FIN_DEFAULT", "Placeholder — replace with your contracts", "USD", "2026-01-01",
             "Example only. Add one scenario per lease/contract."]])
    fr = _sheet(wb, "Financial_Rules",
                ["Scenario_ID", "Scope_Type", "Scope_Value", "Method_Type", "Param_Set_ID"],
                [["FIN_DEFAULT", "CATEGORY", "ENGINE", "MR_FH_FC_RATIO", "P_ENG_MR"],
                 ["FIN_DEFAULT", "ALL", "", "TM_WORKSCOPE", "P_TM"]])
    _add_dropdown(fr, 2, "ScopeType")
    _add_dropdown(fr, 4, "MethodType")
    mr = _sheet(wb, "MR_Rates",
                ["Param_Set_ID", "Component_Scope_Type", "Component_Scope_Value", "Rate_Basis",
                 "Rate", "Escalation_Pct", "Reimbursable", "Reimbursement_Cap"],
                [["P_ENG_MR", "CATEGORY", "ENGINE", "PER_FH", 0, 0.03, "TRUE", ""]])
    _add_dropdown(mr, 2, "ScopeType"); _add_dropdown(mr, 4, "RateBasis"); _add_dropdown(mr, 7, "Bool")
    sc = _sheet(wb, "SV_Cost_Matrix",
                ["Param_Set_ID", "Workscope_Class", "Event_Type", "Fixed_Cost", "Cost_Per_FH", "LLP_Replacement_Cost"],
                [["P_TM", "HEAVY", "HEAVY_SV", 0, 0, 0]])
    _add_dropdown(sc, 2, "WorkscopeClass"); _add_dropdown(sc, 3, "EventType")
    eo = _sheet(wb, "EOL_Terms",
                ["Param_Set_ID", "Component_Scope_Type", "Component_Scope_Value", "Comp_Basis", "Rate"],
                [["P_LLP_EOL", "CATEGORY", "LLP", "PER_CONSUMED_FC", 0]])
    _add_dropdown(eo, 2, "ScopeType"); _add_dropdown(eo, 4, "CompBasis")
    _sheet(wb, "Financial_Params", ["Param_Set_ID", "Param_Name", "Param_Value"],
           [["P_TM", "note", "Fill shop-visit costs here"]])


def _build_instructions(ws):
    ws.column_dimensions["A"].width = 100
    ws.sheet_view.showGridLines = False
    lines = [
        ("Fleet ERP — Relational Database Workbook (filled with your real data)", TITLE_FONT),
        ("This file was generated FROM your DATA INPUT.xlsx into the new unit-based structure. "
         "An aircraft = an airframe + serialized engines + an APU, each unit with its OWN utilisation.", NOTE_FONT),
        ("", None),
        ("Sheets, in order:", Font(bold=True, size=12, color="1F3864")),
        ("• Config — global run settings.", NOTE_FONT),
        ("• Unit_Types — every airframe/engine/APU type + how its LLPs behave.", NOTE_FONT),
        ("• Events — ALL monitored events in one sheet; Unit_Kind column tags each as "
         "AIRFRAME/ENGINE/APU. One row per event; engine/APU events are position-agnostic.", NOTE_FONT),
        ("• Event_Links — the dependency graph: one row per link (an event can link to several). "
         "HARD = always reset target with source; SOFT_BUNDLE = reset target only if it's within "
         "Soft_Threshold_Years of being due (remaining life < years × annual utilisation; "
         "1 = under a year left, 0.5 = under half a year). Rename via Event_Name; keep Event_ID stable.", NOTE_FONT),
        ("• Fleet — one row per aircraft; names its LHENG/RHENG/APU serials + its lease dates.", NOTE_FONT),
        ("• Components — every serialized engine & APU. Installed_On_Tail = the aircraft it flies on; "
         "BLANK = SPARE (still monitored). Position: LHENG / RHENG / APU / SPARE.", NOTE_FONT),
        ("• Unit_Status — current counters (SLSV / TSN / CSN) per unit × event, the snapshot you update each cycle.", NOTE_FONT),
        ("• Utilization — per unit (airframe / each engine / each APU): ANNUAL FH & FC, FH/FC ratio, "
         "evolution, min/max, AOG, and 12 monthly % columns (Jan…Dec) that distribute the annual total "
         "across the year. Monthly FH = Annual_FH × that month's %.", NOTE_FONT),
        ("• LLP_Registry — your company LLP database, one row per LLP per engine serial (ESN).", NOTE_FONT),
        ("• Financial_* — contracts, routing rules, and rate parameters (PLACEHOLDERS — fill with real terms).", NOTE_FONT),
        ("", None),
        ("You name the IDs (tails, ESNs, APU serials, profile names, event IDs). Dropdown fields are locked "
         "to valid values. LLPs come from LLP_Registry — never hand-type them.", NOTE_FONT),
        ("", None),
        ("PLEASE VERIFY: drivers and limits were inferred from the old templates — scan the Events sheet "
         "and correct any Driver/Hard_Limit that looks wrong. Spares were auto-detected; check the Components sheet.", Font(bold=True, color="C00000", size=10)),
    ]
    r = 1
    for text, font in lines:
        c = ws.cell(row=r, column=1, value=text)
        if font:
            c.font = font
        c.alignment = Alignment(wrap_text=True, vertical="top")
        if text == "":
            ws.row_dimensions[r].height = 6
        elif len(text) > 90:
            ws.row_dimensions[r].height = 30
        r += 1


if __name__ == "__main__":
    src = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("/Users/telischotzai/Desktop/DATA INPUT.xlsx")
    llp = Path(sys.argv[2]) if len(sys.argv) > 2 else Path("/Users/telischotzai/Desktop/llps report.xlsx")
    out, stats = build(src, llp)
    print(f"Wrote {out}")
    for k, v in stats.items():
        print(f"  {k}: {v}")
