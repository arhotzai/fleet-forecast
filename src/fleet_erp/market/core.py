"""
Market Intelligence core.

Turns a raw IBA asset export (one row per aircraft) into the lessor tables of the
management "Lessors Update": every aircraft is mapped through the Settings
mappings kept in assets/market_settings.json

    manager name   -> lessor group      (unmapped -> "Other")
    aircraft model -> family            (unmapped -> "Other")
    IBA status     -> Active / On Order / Other
    lessor group   -> funding / main shareholders

and then counted by lessor group x status x family.

Nothing here touches disk except reading the Settings file: the export, the
generated tables and any airline relationship data live in memory only.
"""
from __future__ import annotations

import io
import json
import os
import re
from datetime import date, datetime

import pandas as pd
from openpyxl import load_workbook

_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
_SETTINGS_PATH = os.path.join(_ROOT, "assets", "market_settings.json")

OTHER = "Other"
ACTIVE, ON_ORDER = "Active", "On Order"
_MISSING = {"", "-", "n/a", "na", "none", "nan", "unknown"}
# idle airframes: counted as Active by the Settings, but the ones a lessor can place soonest
_IDLE = {"stored", "parked"}
# natural replacement family when a lease runs off
SUCCESSOR = {"A320ceo": "A320neo", "A330ceo": "A330neo"}

# canonical column -> accepted (normalised) header spellings
_HEADERS = {
    "id": ("aircraftid",),
    "msn": ("aircraftsnfamily", "aircraftsn", "msn", "serialnumber"),
    "reg": ("registration", "reg"),
    "model": ("aircraftmodel", "model"),
    "operator": ("operator",),
    "engine": ("enginemodel", "engine"),
    "build": ("builddate",),
    "delivery": ("originaldelivery", "deliverydate"),
    "status": ("status",),
    "manager": ("manager",),
    "owner": ("owner",),
    "n": ("aircraft", "noaircraft", "count"),
}
_REQUIRED = ("model", "operator", "status", "manager")


# ───────────────────────── helpers ─────────────────────────
def _clean(v) -> str:
    if v is None:
        return ""
    s = " ".join(str(v).split())
    return "" if s.casefold() in _MISSING else s


def _key(v) -> str:
    return _clean(v).casefold()


def _slug(v) -> str:
    return re.sub(r"[^a-z0-9]", "", str(v or "").casefold())


def _i(v) -> int:
    try:
        return int(v)
    except (TypeError, ValueError):
        return 0


def _ref_year(df: pd.DataFrame) -> int:
    """Year the ages are measured from: the newest active airframe, never in the future."""
    now = datetime.now().year
    y = df.loc[df["status2"] == ACTIVE, "year"]
    return int(min(y.max(), now)) if y.notna().any() else now


# Settings tab layout (Lessors Update workbook, copied into the fleet database):
# (key column, value column), zero-based — A:B, D:E, G:H, J:K
SETTINGS_SHEET = "Settings"
SETTINGS_PAIRS = {"managers": (0, 1), "models": (3, 4), "statuses": (6, 7), "funding": (9, 10)}
SETTINGS_HEADER = ["Manager", "Lessor group", None, "Model", "Family", None,
                   "Status", "Counted as", None, "Lessor group", "Funding / main shareholders"]
# family column order of the management summary
FAMILY_ORDER = ["A320ceo", "A320neo", "A321F", "A330ceo", "A330neo", "787",
                "ATR42", "ATR72", "A220", "E2", "Other"]


def _index(cfg: dict) -> dict:
    """Add the casefolded lookups every mapping uses."""
    for k in ("managers", "models", "statuses", "funding"):
        cfg["_" + k] = {_key(a): b for a, b in cfg[k].items()}
    return cfg


def settings_from_rows(rows) -> dict | None:
    """Build the mappings from the Settings tab rows (A:B, D:E, G:H, J:K). A header row
    (\"Manager | Lessor group | …\") is skipped. None when the tab holds no mapping."""
    hdr = {_key(h) for h in SETTINGS_HEADER if h}
    out: dict[str, dict] = {k: {} for k in SETTINGS_PAIRS}
    for row in rows:
        row = list(row or ()) + [None] * 11
        for name, (ki, vi) in SETTINGS_PAIRS.items():
            k, v = _clean(row[ki]), _clean(row[vi])
            if k and v and k not in out[name] and not (_key(k) in hdr and _key(v) in hdr):
                out[name][k] = v
    if not any(out.values()):
        return None
    seen = list(dict.fromkeys(out["models"].values()))
    out["familyOrder"] = ([f for f in FAMILY_ORDER if f in seen or f == OTHER]
                          + [f for f in seen if f not in FAMILY_ORDER])
    return out


def settings_from_workbook(raw: bytes) -> dict | None:
    """The Settings mappings from a workbook's 'Settings' tab, or None if it has none."""
    wb = load_workbook(io.BytesIO(raw), read_only=True, data_only=True)
    try:
        if SETTINGS_SHEET not in wb.sheetnames:
            return None
        cfg = settings_from_rows(wb[SETTINGS_SHEET].iter_rows(values_only=True))
    finally:
        wb.close()
    return _index(cfg) if cfg else None


def load_settings() -> dict:
    """Default mappings (assets/market_settings.json) — used until a fleet database with a
    Settings tab is uploaded."""
    with open(_SETTINGS_PATH, encoding="utf-8") as f:
        return _index(json.load(f))


def apply_settings(df: pd.DataFrame, cfg: dict) -> pd.DataFrame:
    """(Re)map a parsed export's raw model / status / manager through the Settings."""
    def mapped(series, table, default):
        uniq = {v: table.get(v.casefold(), default) for v in series.unique()}
        return series.map(uniq)

    df = df.copy()
    df["family"] = mapped(df["model"], cfg["_models"], OTHER)
    df["status2"] = mapped(df["status"], cfg["_statuses"], OTHER)
    df["group"] = mapped(df["manager"], cfg["_managers"], OTHER)
    return df


def _funding(cfg: dict, group: str) -> tuple[str, str, str]:
    """(full label, country code, shareholder) for a lessor group."""
    full = cfg["_funding"].get(_key(group), "")
    if " - " in full:
        country, holder = full.split(" - ", 1)
        return full, country.strip(), holder.strip()
    return full, "", full


# ───────────────────────── parsing ─────────────────────────
def parse_iba(raw: bytes, cfg: dict) -> pd.DataFrame:
    """Read an IBA asset export into one compact row per aircraft, mapped through Settings."""
    wb = load_workbook(io.BytesIO(raw), read_only=True, data_only=True)
    best = None
    for ws in wb.worksheets:
        rows = ws.iter_rows(values_only=True)
        for i, row in enumerate(rows):
            if i > 20:
                break
            names = [_slug(c) for c in row]
            idx = {}
            for canon, alts in _HEADERS.items():
                for j, nm in enumerate(names):
                    if nm in alts:
                        idx[canon] = j
                        break
            if all(k in idx for k in _REQUIRED):
                best = (idx, list(rows))
                break
        if best:
            break
    wb.close()
    if not best:
        raise ValueError(
            "This does not look like an asset report — expected columns "
            "'Aircraft Model', 'Operator', 'Status' and 'Manager'.")
    idx, data = best
    width = max(idx.values()) + 1

    def col(name):
        j = idx.get(name)
        if j is None:
            return [None] * len(data)
        return [r[j] if len(r) > j else None for r in data]

    data = [r for r in data if r and len(r) >= 1 and any(c is not None for c in r[:width])]
    df = pd.DataFrame({k: col(k) for k in _HEADERS})

    def text(series):
        uniq = {v: _clean(v) for v in series.unique()}
        return series.map(uniq)

    for c in ("msn", "reg", "model", "operator", "engine", "status", "manager", "owner"):
        df[c] = text(df[c].astype(object).where(df[c].notna(), ""))
    df = df[(df["model"] != "") | (df["operator"] != "") | (df["status"] != "")].reset_index(drop=True)
    if df.empty:
        raise ValueError("The export has no aircraft rows.")

    n = pd.to_numeric(df["n"], errors="coerce") if "n" in idx else None
    df["n"] = (n.fillna(1) if n is not None else 1)
    df["n"] = df["n"].astype(int)

    def year(series):
        def one(v):
            if isinstance(v, (datetime, date)):
                return v.year
            if isinstance(v, str) and len(v.strip()) >= 6:
                ts = pd.to_datetime(v.strip(), errors="coerce")
                return None if pd.isna(ts) else ts.year
            return None
        y = pd.to_numeric(pd.Series([one(v) for v in series], index=series.index), errors="coerce")
        return y.where((y >= 1940) & (y <= 2040))     # undated orders carry a far-future placeholder

    df["year"] = year(df["build"])
    df["delYear"] = year(df["delivery"])
    df = df.drop(columns=["build", "delivery", "id"])

    df = apply_settings(df, cfg)
    df["idle"] = df["status"].str.casefold().isin(_IDLE)
    df["ownerOp"] = (df["operator"] != "") & (df["operator"].str.casefold() == df["owner"].str.casefold())
    return df


# ───────────────────────── aggregation ─────────────────────────
def _families(cfg: dict, df: pd.DataFrame) -> list[str]:
    seen = set(df["family"].unique())
    fams = [f for f in cfg["familyOrder"] if f in seen and f != OTHER]
    fams += sorted(f for f in seen if f not in cfg["familyOrder"] and f != OTHER)
    return fams + [OTHER]


def _pivot(df: pd.DataFrame, status2: str) -> pd.DataFrame:
    d = df[df["status2"] == status2]
    if d.empty:
        return pd.DataFrame()
    return d.pivot_table(index="group", columns="family", values="n", aggfunc="sum", fill_value=0)


def _row_counts(p: pd.DataFrame, group: str, fams: list[str]) -> dict:
    if p.empty or group not in p.index:
        return {}
    r = p.loc[group]
    return {f: int(r[f]) for f in fams if f in p.columns and r[f]}


def _sum(d: dict) -> int:
    return int(sum(d.values()))


def _add(a: dict, b: dict) -> dict:
    out = dict(a)
    for k, v in b.items():
        out[k] = out.get(k, 0) + v
    return out


def report(df: pd.DataFrame, cfg: dict, prev: pd.DataFrame | None = None,
           relationship: dict | None = None) -> dict:
    """The lessor league (management "Summary 2") plus market cuts around it."""
    fams = _families(cfg, df)
    act, orders = _pivot(df, ACTIVE), _pivot(df, ON_ORDER)
    prev_act = _pivot(prev, ACTIVE) if prev is not None else None
    ref_year = _ref_year(df)

    active_df = df[df["status2"] == ACTIVE]
    age = active_df.dropna(subset=["year"]).assign(age=lambda d: ref_year - d["year"])
    avg_age = age.groupby("group")["age"].mean().round(1).to_dict()
    idle = active_df[active_df["idle"]].groupby("group")["n"].sum().to_dict()
    customers = active_df[active_df["operator"] != ""].groupby("group")["operator"].nunique().to_dict()

    rel_by = (relationship or {}).get("byLessor", {})
    grand_active = int(active_df["n"].sum())
    grand_orders = int(df.loc[df["status2"] == ON_ORDER, "n"].sum())

    def build(group: str) -> dict:
        a, o = _row_counts(act, group, fams), _row_counts(orders, group, fams)
        at, ot = _sum(a), _sum(o)
        full, country, holder = _funding(cfg, group)
        row = {
            "name": group, "funding": full, "country": country, "shareholder": holder,
            "active": a, "activeTotal": at, "orders": o, "ordersTotal": ot,
            "share": round(at / grand_active * 100, 2) if grand_active else 0,
            "orderShare": round(ot / grand_orders * 100, 2) if grand_orders else 0,
            "backlog": round(ot / at * 100, 1) if at else None,
            "avgAge": avg_age.get(group), "idle": _i(idle.get(group, 0)),
            "customers": _i(customers.get(group, 0)),
            "prevTotal": None, "delta": None,
            "rel": rel_by.get(group, {}), "relTotal": _sum(rel_by.get(group, {})),
        }
        if prev_act is not None:
            pt = _sum(_row_counts(prev_act, group, list(prev_act.columns)))
            row["prevTotal"], row["delta"] = pt, at - pt
        return row

    groups = sorted(set(act.index) | set(orders.index) | set(rel_by))
    lessors = [build(g) for g in groups if g != OTHER]
    lessors = [r for r in lessors if r["activeTotal"] or r["ordersTotal"] or r["relTotal"]]
    lessors.sort(key=lambda r: (-r["activeTotal"], -r["ordersTotal"], r["name"]))
    for i, r in enumerate(lessors, 1):
        r["rank"] = i
    other = build(OTHER)
    other["name"] = "Other (airlines & unlisted managers)"

    def total(rows: list[dict], name: str) -> dict:
        a, o, rel = {}, {}, {}
        for r in rows:
            a, o, rel = _add(a, r["active"]), _add(o, r["orders"]), _add(rel, r["rel"])
        at, ot = _sum(a), _sum(o)
        t = {"name": name, "funding": "", "country": "", "shareholder": "",
             "active": a, "activeTotal": at, "orders": o, "ordersTotal": ot,
             "share": round(at / grand_active * 100, 2) if grand_active else 0,
             "orderShare": round(ot / grand_orders * 100, 2) if grand_orders else 0,
             "backlog": round(ot / at * 100, 1) if at else None,
             "avgAge": None, "idle": sum(r["idle"] for r in rows),
             "customers": None, "prevTotal": None, "delta": None,
             "rel": rel, "relTotal": _sum(rel)}
        if prev_act is not None:
            t["prevTotal"] = sum(r["prevTotal"] or 0 for r in rows)
            t["delta"] = at - t["prevTotal"]
        return t

    lessor_total = total(lessors, "LESSORS TOTAL")
    grand_total = total(lessors + [other], "GRAND TOTAL")

    # family view: how much of each family the listed lessors manage
    fam_rows = []
    for f in fams:
        a = int(act[f].sum()) if (not act.empty and f in act.columns) else 0
        o = int(orders[f].sum()) if (not orders.empty and f in orders.columns) else 0
        la, lo = lessor_total["active"].get(f, 0), lessor_total["orders"].get(f, 0)
        fa = age[age["family"] == f]
        top = max(lessors, key=lambda r: r["active"].get(f, 0), default=None)
        top_o = max(lessors, key=lambda r: r["orders"].get(f, 0), default=None)
        fam_rows.append({
            "family": f, "active": a, "onOrder": o, "lessorActive": la, "lessorOrders": lo,
            "lessorShare": round(la / a * 100, 1) if a else None,
            "lessorOrderShare": round(lo / o * 100, 1) if o else None,
            "backlog": round(o / a * 100, 1) if a else None,
            "avgAge": round(float(fa["age"].mean()), 1) if len(fa) else None,
            "idle": int(active_df.loc[active_df["idle"] & (active_df["family"] == f), "n"].sum()),
            "topLessor": top["name"] if top and top["active"].get(f) else "",
            "topLessorCount": top["active"].get(f, 0) if top else 0,
            "topOrderLessor": top_o["name"] if top_o and top_o["orders"].get(f) else "",
            "topOrderCount": top_o["orders"].get(f, 0) if top_o else 0,
        })

    # capital origin: where the money behind the listed lessors comes from
    origin: dict[str, dict] = {}
    for r in lessors:
        c = r["country"] or "Not set"
        o = origin.setdefault(c, {"country": c, "lessors": 0, "active": 0, "orders": 0, "names": []})
        o["lessors"] += 1
        o["active"] += r["activeTotal"]
        o["orders"] += r["ordersTotal"]
        o["names"].append(r["name"])
    origins = sorted(origin.values(), key=lambda o: -o["active"])
    for o in origins:
        o["share"] = round(o["active"] / lessor_total["activeTotal"] * 100, 1) if lessor_total["activeTotal"] else 0
        o["names"] = o["names"][:6]

    # data quality: what the Settings do not cover yet
    un = df[(df["group"] == OTHER) & (df["manager"] != "") & df["status2"].isin([ACTIVE, ON_ORDER])]
    third = un[un["manager"].str.casefold() != un["operator"].str.casefold()]
    um = (third.pivot_table(index="manager", columns="status2", values="n", aggfunc="sum", fill_value=0)
          if not third.empty else pd.DataFrame())
    ops = third.groupby("manager")["operator"].nunique() if not third.empty else {}
    unmapped_managers = []
    for name, r in um.iterrows():
        a, o = _i(r.get(ACTIVE, 0)), _i(r.get(ON_ORDER, 0))
        unmapped_managers.append({"name": name, "active": a, "orders": o, "operators": _i(ops.get(name, 0))})
    unmapped_managers.sort(key=lambda r: -(r["active"] + r["orders"]))
    um2 = df[(df["family"] == OTHER) & df["status2"].isin([ACTIVE, ON_ORDER])]
    mm = (um2.pivot_table(index="model", columns="status2", values="n", aggfunc="sum", fill_value=0)
          if not um2.empty else pd.DataFrame())
    unmapped_models = sorted(
        ({"model": m or "(blank)", "active": _i(r.get(ACTIVE, 0)), "orders": _i(r.get(ON_ORDER, 0))}
         for m, r in mm.iterrows()), key=lambda r: -(r["active"] + r["orders"]))
    no_funding = [r["name"] for r in lessors if not r["funding"]]
    unknown_status = sorted(s for s in df["status"].unique() if s and s.casefold() not in cfg["_statuses"])

    top5 = sum(r["activeTotal"] for r in lessors[:5])
    by_orders = sorted(lessors, key=lambda r: -r["ordersTotal"])
    status_counts = df.groupby("status")["n"].sum().sort_values(ascending=False)
    owned = active_df[active_df["ownerOp"]]["n"].sum()

    return {
        "families": fams,
        "orderFamilies": [f for f in fams if grand_total["orders"].get(f)],
        "relTypes": (relationship or {}).get("types", []),
        "relationship": {k: v for k, v in (relationship or {}).items() if k != "byLessor"} or None,
        "hasPrev": prev is not None,
        "kpis": {
            "rows": int(df["n"].sum()),
            "active": grand_active, "onOrder": grand_orders,
            "lessorActive": lessor_total["activeTotal"], "lessorOrders": lessor_total["ordersTotal"],
            "lessorShare": round(lessor_total["activeTotal"] / grand_active * 100, 1) if grand_active else 0,
            "lessorOrderShare": round(lessor_total["ordersTotal"] / grand_orders * 100, 1) if grand_orders else 0,
            "lessors": len(lessors),
            "top5Share": round(top5 / lessor_total["activeTotal"] * 100, 1) if lessor_total["activeTotal"] else 0,
            "biggestLessor": lessors[0]["name"] if lessors else "",
            "biggestLessorFleet": lessors[0]["activeTotal"] if lessors else 0,
            "biggestOrderbook": by_orders[0]["name"] if by_orders else "",
            "biggestOrderbookCount": by_orders[0]["ordersTotal"] if by_orders else 0,
            "idle": int(active_df.loc[active_df["idle"], "n"].sum()),
            "lessorIdle": lessor_total["idle"],
            "operatorOwnedShare": round(float(owned) / grand_active * 100, 1) if grand_active else 0,
            "refYear": ref_year,
            "operators": int(active_df.loc[active_df["operator"] != "", "operator"].nunique()),
        },
        "lessors": lessors, "other": other, "lessorTotal": lessor_total, "total": grand_total,
        "familyTable": fam_rows,
        "origins": origins,
        "statuses": [{"status": s, "mapped": cfg["_statuses"].get(s.casefold(), OTHER), "count": int(c)}
                     for s, c in status_counts.items()],
        "quality": {
            "unmappedManagers": unmapped_managers[:60], "unmappedManagerCount": len(unmapped_managers),
            "unmappedModels": unmapped_models[:60], "unmappedModelCount": len(unmapped_models),
            "noFunding": no_funding, "unknownStatuses": unknown_status,
        },
    }


# ───────────────────────── drill-downs ─────────────────────────
def operators(df: pd.DataFrame, q: str = "", limit: int = 40) -> list[dict]:
    d = df[(df["operator"] != "") & df["status2"].isin([ACTIVE, ON_ORDER])]
    if q:
        d = d[d["operator"].str.casefold().str.contains(q.casefold(), regex=False)]
    if d.empty:
        return []
    p = d.pivot_table(index="operator", columns="status2", values="n", aggfunc="sum", fill_value=0)
    for c in (ACTIVE, ON_ORDER):
        if c not in p.columns:
            p[c] = 0
    p = p.sort_values([ACTIVE, ON_ORDER], ascending=False).head(limit)
    # directory extras for the operators shown: idle, lessor-managed share, lessors used, main family
    act = d[(d["status2"] == ACTIVE) & d["operator"].isin(p.index)]
    idle = act[act["idle"]].groupby("operator")["n"].sum()
    leased = act[act["group"] != OTHER].groupby("operator")["n"].sum()
    nles = act[act["group"] != OTHER].groupby("operator")["group"].nunique()
    # main type: the Settings family, or the raw model when it falls under "Other" (737s, widebodies…)
    lab = act["family"].where(act["family"] != OTHER, act["model"])
    fam = act.assign(lab=lab).groupby(["operator", "lab"])["n"].sum().reset_index().sort_values("n", ascending=False)
    top_fam = fam.drop_duplicates("operator").set_index("operator")["lab"]
    out = []
    for n, r in p.iterrows():
        a_n = int(r[ACTIVE])
        out.append({"name": n, "active": a_n, "orders": int(r[ON_ORDER]),
                    "idle": int(idle.get(n, 0)), "lessors": int(nles.get(n, 0)),
                    "leasedShare": round(float(leased.get(n, 0)) / a_n * 100, 1) if a_n else None,
                    "mainFamily": top_fam.get(n, "")})
    return out


def _aircraft_rows(d: pd.DataFrame, limit: int = 600) -> list[dict]:
    d = d.sort_values(["status2", "family", "year"], na_position="last").head(limit)
    return [{
        "reg": r.reg, "msn": r.msn, "model": r.model, "family": r.family, "status": r.status,
        "status2": r.status2, "operator": r.operator, "manager": r.manager, "owner": r.owner,
        "lessor": r.group, "engine": r.engine,
        "year": int(r.year) if pd.notna(r.year) else None,
    } for r in d.itertuples(index=False)]


def _age_buckets(d: pd.DataFrame, ref_year: int) -> list[dict]:
    a = d.dropna(subset=["year"])
    ages = ref_year - a["year"]
    out = []
    for label, lo, hi in (("0–5", 0, 5), ("6–10", 6, 10), ("11–15", 11, 15), ("16–20", 16, 20), ("21+", 21, 999)):
        out.append({"bucket": label, "count": int(a.loc[(ages >= lo) & (ages <= hi), "n"].sum())})
    return out


def _fam_status(d: pd.DataFrame, fams: list[str]) -> list[dict]:
    out = []
    for f in fams:
        x = d[d["family"] == f]
        a = int(x.loc[x["status2"] == ACTIVE, "n"].sum())
        o = int(x.loc[x["status2"] == ON_ORDER, "n"].sum())
        if a or o:
            out.append({"family": f, "active": a, "onOrder": o,
                        "idle": int(x.loc[x["idle"], "n"].sum())})
    return out


def operator_view(df: pd.DataFrame, cfg: dict, name: str) -> dict | None:
    """Everything the export says about one operator (active fleet + orders)."""
    d = df[(df["operator"].str.casefold() == name.casefold()) & df["status2"].isin([ACTIVE, ON_ORDER])]
    if d.empty:
        return None
    fams = _families(cfg, df)
    ref_year = _ref_year(df)
    act = d[d["status2"] == ACTIVE]
    # exposure: who manages the active fleet (own aircraft collapse to "Owned / own-managed")
    label = act["group"].where(act["group"] != OTHER, act["manager"])
    own = (act["manager"] == "") | (act["manager"].str.casefold() == name.casefold())
    label = label.where(~own, "Owned / own-managed")
    exp = act.assign(label=label).pivot_table(index="label", columns="family", values="n",
                                              aggfunc="sum", fill_value=0) if not act.empty else pd.DataFrame()
    exposure = []
    for lab, r in exp.iterrows():
        by = {f: int(r[f]) for f in exp.columns if r[f]}
        full, country, _ = _funding(cfg, lab)
        exposure.append({"lessor": lab, "funding": full, "country": country,
                         "listed": _key(lab) in {_key(v) for v in cfg["managers"].values()},
                         "byFamily": by, "total": _sum(by)})
    exposure.sort(key=lambda r: -r["total"])
    total = int(act["n"].sum())
    ages = (ref_year - act.dropna(subset=["year"])["year"])
    return {
        "name": str(d["operator"].iloc[0]),
        "active": total, "orders": int(d.loc[d["status2"] == ON_ORDER, "n"].sum()),
        "idle": int(act.loc[act["idle"], "n"].sum()),
        "avgAge": round(float(ages.mean()), 1) if len(ages) else None,
        "leasedShare": round(float(act.loc[~own, "n"].sum()) / total * 100, 1) if total else None,
        "families": [f for f in fams if (d["family"] == f).any()],
        "byFamily": _fam_status(d, fams), "exposure": exposure,
        "ageProfile": _age_buckets(act, ref_year),
        "aircraft": _aircraft_rows(d),
    }


def lessor_view(df: pd.DataFrame, cfg: dict, group: str) -> dict | None:
    """Profile of one lessor group: who flies its aircraft, how old, what is coming."""
    d = df[(df["group"].str.casefold() == group.casefold()) & df["status2"].isin([ACTIVE, ON_ORDER])]
    if d.empty:
        return None
    fams = _families(cfg, df)
    ref_year = _ref_year(df)
    act, od = d[d["status2"] == ACTIVE], d[d["status2"] == ON_ORDER]
    full, country, holder = _funding(cfg, group)

    def top_ops(x: pd.DataFrame, k: int = 15) -> list[dict]:
        x = x[x["operator"] != ""]
        if x.empty:
            return []
        p = x.pivot_table(index="operator", columns="family", values="n", aggfunc="sum", fill_value=0)
        p["__t"] = p.sum(axis=1)
        p = p.sort_values("__t", ascending=False).head(k)
        return [{"operator": n, "total": int(r["__t"]),
                 "byFamily": {f: int(r[f]) for f in p.columns if f != "__t" and r[f]}}
                for n, r in p.iterrows()]

    mgr = d.groupby("manager")["n"].sum().sort_values(ascending=False)
    ages = ref_year - act.dropna(subset=["year"])["year"]
    idle = act[act["idle"]]
    return {
        "name": str(d["group"].iloc[0]), "funding": full, "country": country, "shareholder": holder,
        "active": int(act["n"].sum()), "orders": int(od["n"].sum()),
        "idle": int(idle["n"].sum()),
        "avgAge": round(float(ages.mean()), 1) if len(ages) else None,
        "customers": int(act.loc[act["operator"] != "", "operator"].nunique()),
        "byFamily": _fam_status(d, fams),
        "ageProfile": _age_buckets(act, ref_year),
        "topOperators": top_ops(act), "orderCustomers": top_ops(od, 10),
        "managers": [{"name": n, "count": int(c)} for n, c in mgr.items()],
        "idleAircraft": _aircraft_rows(idle, 200),
    }


# ───────────────────────── airline relationship (runtime only) ─────────────────────────
def _match_group(name: str, cfg: dict, groups: dict[str, str]) -> str | None:
    """Map a lessor name from the fleet database to a Settings lessor group."""
    k = _key(name)
    if not k:
        return None
    if k in cfg["_managers"]:
        return cfg["_managers"][k]
    s = _slug(name)
    if s in groups:
        return groups[s]
    best = None
    for gs, g in groups.items():
        if len(gs) >= 4 and (gs in s or (len(s) >= 4 and s in gs)):
            if best is None or len(gs) > len(_slug(best)):
                best = g
    if best:
        return best
    for mk, g in cfg["_managers"].items():
        ms = _slug(mk)
        if len(ms) >= 6 and (ms in s or (len(s) >= 6 and s in ms)):
            return g
    return None


def relationship_from_fleet(aircraft: list[dict], cfg: dict, airline: str = "") -> dict | None:
    """Current relationship by lessor x type, from the loaded fleet database (never stored)."""
    rows = [a for a in aircraft if a.get("active")]
    if not rows:
        return None
    groups = {_slug(g): g for g in set(cfg["managers"].values())}
    by_lessor: dict[str, dict] = {}
    unmatched: dict[str, dict] = {}
    owned: dict[str, int] = {}
    types: dict[str, int] = {}
    expiries: dict[str, dict] = {}
    for a in rows:
        t = str(a.get("type") or "").strip() or "Unknown"
        t = cfg["_models"].get(t.casefold(), t)
        types[t] = types.get(t, 0) + 1
        leased = str(a.get("leasedOwned") or "").strip().casefold().startswith("leas")
        if not leased:
            owned[t] = owned.get(t, 0) + 1
            continue
        end = a.get("extensionDate") or a.get("redeliveryDate")
        if end:
            y = str(end)[:4]
            e = expiries.setdefault(y, {})
            e[t] = e.get(t, 0) + 1
        g = _match_group(a.get("lessor") or "", cfg, groups)
        tgt = by_lessor.setdefault(g, {}) if g else unmatched.setdefault(
            str(a.get("lessor") or "Unknown lessor"), {})
        tgt[t] = tgt.get(t, 0) + 1
    order = cfg["familyOrder"]
    tlist = sorted(types, key=lambda t: (order.index(t) if t in order else 99, t))
    return {
        "source": "fleet", "airline": airline, "types": tlist, "composition": types,
        "byLessor": by_lessor,
        "unmatched": [{"lessor": k, "byType": v, "total": _sum(v)} for k, v in
                      sorted(unmatched.items(), key=lambda kv: -_sum(kv[1]))],
        "owned": owned, "ownedTotal": _sum(owned),
        "leasedTotal": sum(_sum(v) for v in by_lessor.values()) + sum(_sum(v) for v in unmatched.values()),
        "total": len(rows),
        "expiries": [{"year": y, "byType": v, "total": _sum(v)} for y, v in sorted(expiries.items())],
    }


def relationship_from_operator(df: pd.DataFrame, cfg: dict, name: str) -> dict | None:
    """Same relationship view for any operator in the export (as IBA reports it)."""
    act = df[(df["operator"].str.casefold() == name.casefold()) & (df["status2"] == ACTIVE)]
    if act.empty:
        return None
    by_lessor: dict[str, dict] = {}
    unmatched: dict[str, dict] = {}
    owned: dict[str, int] = {}
    types: dict[str, int] = {}
    for r in act.itertuples(index=False):
        t = r.family
        types[t] = types.get(t, 0) + r.n
        if r.manager == "" or r.manager.casefold() == name.casefold():
            owned[t] = owned.get(t, 0) + r.n
        elif r.group != OTHER:
            g = by_lessor.setdefault(r.group, {})
            g[t] = g.get(t, 0) + r.n
        else:
            g = unmatched.setdefault(r.manager, {})
            g[t] = g.get(t, 0) + r.n
    order = cfg["familyOrder"]
    return {
        "source": "iba", "airline": str(act["operator"].iloc[0]),
        "types": sorted(types, key=lambda t: (order.index(t) if t in order else 99, t)),
        "composition": types, "byLessor": by_lessor,
        "unmatched": [{"lessor": k, "byType": v, "total": _sum(v)} for k, v in
                      sorted(unmatched.items(), key=lambda kv: -_sum(kv[1]))],
        "owned": owned, "ownedTotal": _sum(owned),
        "leasedTotal": sum(_sum(v) for v in by_lessor.values()) + sum(_sum(v) for v in unmatched.values()),
        "total": int(act["n"].sum()), "expiries": [],
    }


# ───────────────────────── fleet fit ─────────────────────────
def fleet_fit(df: pd.DataFrame, cfg: dict, rep: dict, relationship: dict) -> dict:
    """Where the airline's fleet sits in the lessor market, and who could fill the gaps."""
    comp = relationship.get("composition", {})
    total = sum(comp.values()) or 1
    existing = set(relationship.get("byLessor", {}))
    market = {r["family"]: r for r in rep["familyTable"]}
    lessors = rep["lessors"]

    targets: list[str] = []
    for f in comp:
        for t in (f, SUCCESSOR.get(f)):
            if t and t in market and t != OTHER and t not in targets:
                targets.append(t)

    families = []
    for f in sorted(set(comp) | set(targets), key=lambda x: -comp.get(x, 0)):
        m = market.get(f)
        holders = sorted(
            ({"lessor": r["name"], "active": r["active"].get(f, 0), "orders": r["orders"].get(f, 0),
              "existing": r["name"] in existing} for r in lessors
             if r["active"].get(f, 0) or r["orders"].get(f, 0)),
            key=lambda x: -(x["active"] + x["orders"]))[:8]
        families.append({
            "family": f, "airline": comp.get(f, 0),
            "airlineShare": round(comp.get(f, 0) / total * 100, 1),
            "successor": SUCCESSOR.get(f, ""),
            "inMarket": bool(m),
            "marketActive": m["active"] if m else 0, "marketOrders": m["onOrder"] if m else 0,
            "lessorActive": m["lessorActive"] if m else 0, "lessorOrders": m["lessorOrders"] if m else 0,
            "lessorShare": m["lessorShare"] if m else None,
            "marketAvgAge": m["avgAge"] if m else None,
            "idle": m["idle"] if m else 0,
            "topLessors": holders,
        })

    # sourcing shortlist: lessors ranked by what they hold / have coming in the target families
    idle_df = df[(df["status2"] == ACTIVE) & df["idle"] & df["family"].isin(targets) & (df["group"] != OTHER)]
    idle_by = idle_df.groupby("group")["n"].sum().to_dict() if not idle_df.empty else {}
    shortlist = []
    for r in lessors:
        a = {f: r["active"].get(f, 0) for f in targets if r["active"].get(f, 0)}
        o = {f: r["orders"].get(f, 0) for f in targets if r["orders"].get(f, 0)}
        if not a and not o:
            continue
        idle_n = _i(idle_by.get(r["name"], 0))
        shortlist.append({
            "lessor": r["name"], "funding": r["funding"], "existing": r["name"] in existing,
            "relTotal": r["relTotal"], "active": a, "activeTotal": _sum(a),
            "orders": o, "ordersTotal": _sum(o), "idle": idle_n,
            # near-term supply (orders + idle airframes) weighs more than installed base
            "score": round(_sum(o) * 1.0 + idle_n * 1.5 + _sum(a) * 0.15, 1),
        })
    shortlist.sort(key=lambda r: -r["score"])

    rel_rows = sorted(({"lessor": k, "total": _sum(v)} for k, v in relationship.get("byLessor", {}).items()),
                      key=lambda r: -r["total"])
    rel_rows += [{"lessor": u["lessor"], "total": u["total"]} for u in relationship.get("unmatched", [])]
    leased = sum(r["total"] for r in rel_rows) or 1
    hhi = round(sum((r["total"] / leased * 100) ** 2 for r in rel_rows))
    return {
        "airline": relationship.get("airline", ""), "source": relationship.get("source"),
        "total": relationship.get("total", 0),
        "ownedTotal": relationship.get("ownedTotal", 0), "leasedTotal": relationship.get("leasedTotal", 0),
        "targets": targets, "families": families,
        "shortlist": shortlist[:25],
        "untapped": [s for s in shortlist if not s["existing"]][:10],
        "concentration": {
            "lessors": len(rel_rows), "hhi": hhi if rel_rows else None,
            "top": rel_rows[0]["lessor"] if rel_rows else "",
            "topShare": round(rel_rows[0]["total"] / leased * 100, 1) if rel_rows else None,
            "rows": [{**r, "share": round(r["total"] / leased * 100, 1)} for r in rel_rows],
        },
        "expiries": relationship.get("expiries", []),
    }


# ───────────────────────── aircraft-type view ─────────────────────────
def types(df: pd.DataFrame) -> list[dict]:
    """Families (Settings) with their models, active + on-order counts — for the type picker."""
    d = df[df["status2"].isin([ACTIVE, ON_ORDER])]
    g = d.groupby(["family", "model", "status2"])["n"].sum().unstack(fill_value=0)
    for c in (ACTIVE, ON_ORDER):
        if c not in g.columns:
            g[c] = 0
    out: dict[str, dict] = {}
    for (fam, model), r in g.iterrows():
        if not model:
            continue
        f = out.setdefault(fam, {"family": fam, "active": 0, "orders": 0, "models": []})
        a, o = int(r[ACTIVE]), int(r[ON_ORDER])
        f["active"] += a
        f["orders"] += o
        f["models"].append({"model": model, "active": a, "orders": o})
    for f in out.values():
        f["models"].sort(key=lambda m: -(m["active"] + m["orders"]))
    return sorted(out.values(), key=lambda f: (f["family"] == OTHER, -(f["active"] + f["orders"])))


def type_view(df: pd.DataFrame, cfg: dict, family: str, models: list[str] | None = None) -> dict | None:
    """Everything the type tab filters client-side, for one family: an aggregated cube of
    lessor x model x status x idle x build year x delivery year (operator-managed and
    unmapped managers fold into "Other"), the same by operator, and lessor metadata."""
    d = df[(df["family"].str.casefold() == family.casefold()) & df["status2"].isin([ACTIVE, ON_ORDER])]
    all_models = sorted(d["model"].unique().tolist())
    if models:
        want = {m.casefold() for m in models}
        d = d[d["model"].str.casefold().isin(want)]
    if d.empty:
        return None

    def _yr(v):
        return None if pd.isna(v) else int(v)

    keys = ["group", "model", "status2", "idle", "year", "delYear"]
    cube = d.groupby(keys, dropna=False)["n"].sum().reset_index()
    rows = [{"lessor": r.group, "model": r.model, "status": r.status2, "idle": bool(r.idle),
             "year": _yr(r.year), "delYear": _yr(r.delYear), "n": int(r.n)} for r in cube.itertuples()]
    ops = d.groupby(["operator", "group", "model", "status2"])["n"].sum().reset_index()
    op_rows = [{"operator": r.operator or "(no operator)", "lessor": r.group, "model": r.model,
                "status": r.status2, "n": int(r.n)} for r in ops.itertuples()]
    meta = {}
    for g in d["group"].unique():
        if g == OTHER:
            continue
        full, country, holder = _funding(cfg, g)
        meta[g] = {"funding": full, "country": country, "shareholder": holder}
    return {"family": d["family"].iloc[0], "models": all_models,
            "refYear": _ref_year(df), "cube": rows, "operators": op_rows, "lessors": meta}
