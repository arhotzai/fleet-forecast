"""
Ingestion — read the master database workbook + the raw LLP report, merge, validate.

This is the ONLY module that knows sheet/column names. Everything downstream works
on the clean `Dataset` it returns.
"""

from __future__ import annotations

import io
from dataclasses import dataclass, field

import pandas as pd

# Expected sheets in the master database workbook
DB_SHEETS = [
    "Config", "Unit_Types", "Events", "Event_Links",
    "Fleet", "Components", "Unit_Status", "Utilization",
    "LLP_Registry", "Financial_Scenarios", "Financial_Rules",
    "MR_Rates", "SV_Cost_Matrix", "EOL_Terms", "Financial_Params",
]


@dataclass
class Dataset:
    sheets: dict[str, pd.DataFrame]
    issues: list[tuple[str, str]] = field(default_factory=list)   # (level, message)

    def get(self, name: str) -> pd.DataFrame:
        return self.sheets.get(name, pd.DataFrame())

    @property
    def errors(self) -> list[str]:
        return [m for lvl, m in self.issues if lvl == "ERROR"]

    @property
    def warnings(self) -> list[str]:
        return [m for lvl, m in self.issues if lvl == "WARNING"]


def _norm_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    # strip whitespace from object cells
    for c in df.columns:
        if df[c].dtype == object:
            df[c] = df[c].map(lambda v: v.strip() if isinstance(v, str) else v)
    return df


def read_database(file_bytes: bytes) -> dict[str, pd.DataFrame]:
    xl = pd.ExcelFile(io.BytesIO(file_bytes))
    out: dict[str, pd.DataFrame] = {}
    for sheet in xl.sheet_names:
        if sheet.startswith("_"):
            continue
        out[sheet] = _norm_cols(pd.read_excel(xl, sheet))
    return out


def parse_llp_report(file_bytes: bytes) -> pd.DataFrame:
    """
    Parse the raw LLP report exported from the company database.

    Expected 8 columns (duplicate headers in the export):
        engine partno, engine serial, LLP partno, LLP serial,
        life_limit, togo, counter, expected_date
    Computes SLSV (cycles consumed) = life_limit - togo.
    """
    raw = pd.read_excel(io.BytesIO(file_bytes), header=0)
    if raw.shape[1] < 6:
        raise ValueError(
            f"LLP report has {raw.shape[1]} columns; expected at least 6 "
            "(engine p/n, engine s/n, LLP p/n, LLP s/n, life_limit, togo, …)."
        )
    cols = ["eng_partno", "eng_esn", "llp_partno", "llp_sn",
            "life_limit", "togo", "counter", "expected_date"]
    raw = raw.iloc[:, :8]
    raw.columns = cols[:raw.shape[1]]
    for c in ("eng_esn", "llp_partno", "llp_sn", "counter"):
        if c in raw.columns:
            raw[c] = raw[c].map(lambda v: str(v).strip() if pd.notna(v) else None)
    raw["life_limit"] = pd.to_numeric(raw["life_limit"], errors="coerce")
    raw["togo"] = pd.to_numeric(raw["togo"], errors="coerce")

    out = pd.DataFrame({
        "Unit_ID": raw["eng_esn"],
        "LLP_PN": raw.get("llp_partno"),
        "LLP_SN": raw.get("llp_sn"),
        "Counter": raw.get("counter"),
        "Limit_Cycles": raw["life_limit"],
        "ToGo": raw["togo"],
        "SLSV": (raw["life_limit"] - raw["togo"]).round(2),
        "Expected_Date": pd.to_datetime(raw.get("expected_date"), dayfirst=True, errors="coerce"),
    })
    out = out[out["Unit_ID"].notna() & (out["Unit_ID"] != "nan")].reset_index(drop=True)
    return out


def load(db_bytes: bytes, llp_bytes: bytes | None = None) -> Dataset:
    """Load the database and (optionally) overlay a fresh LLP report onto LLP_Registry."""
    sheets = read_database(db_bytes)
    ds = Dataset(sheets=sheets)

    missing = [s for s in ("Fleet", "Components", "Utilization") if s not in sheets or sheets[s].empty]
    for s in missing:
        ds.issues.append(("ERROR", f"Required sheet '{s}' is missing or empty."))

    if llp_bytes:
        try:
            report = parse_llp_report(llp_bytes)
            # carry part descriptions from the existing registry, matched by PART
            # NUMBER (a description belongs to the part number, not the serial).
            existing = sheets.get("LLP_Registry", pd.DataFrame())
            if not existing.empty and {"LLP_PN", "Description"} <= set(existing.columns):
                desc = (existing[["LLP_PN", "Description"]].dropna(subset=["Description"])
                        .drop_duplicates("LLP_PN"))
                report = report.merge(desc, on="LLP_PN", how="left")
            sheets["LLP_Registry"] = report
            ds.issues.append(("WARNING", f"LLP report applied: {len(report)} LLP rows "
                                         f"across {report['Unit_ID'].nunique()} serials "
                                         f"(SLSV computed as life_limit − togo)."))
        except Exception as exc:  # surface a friendly message instead of crashing
            ds.issues.append(("ERROR", f"Could not read the LLP report: {exc}"))

    _cross_check(ds)
    return ds


def _cross_check(ds: Dataset) -> None:
    """Non-fatal referential checks that produce friendly warnings."""
    fleet = ds.get("Fleet")
    units = ds.get("Components")
    util = ds.get("Utilization")

    util_ids = set(util.get("Unit_ID", pd.Series(dtype=str)).astype(str))
    id_col = "Component_ID" if "Component_ID" in units.columns else "Unit_ID"
    unit_ids = set(units.get(id_col, pd.Series(dtype=str)).astype(str))

    # aircraft missing a utilisation row
    if not fleet.empty and "Tail" in fleet.columns:
        miss = [t for t in fleet["Tail"].dropna().astype(str) if t not in util_ids]
        if miss:
            ds.issues.append(("WARNING", f"{len(miss)} aircraft have no Utilization row "
                                         f"(FH/FC events can't be projected): {', '.join(miss[:5])}"
                                         + ("…" if len(miss) > 5 else "")))

    # engine/APU serials referenced by the fleet but absent from Components
    if not fleet.empty:
        for col in ("LHENG_ESN", "RHENG_ESN", "APU_SN"):
            if col in fleet.columns:
                refs = fleet[col].dropna().astype(str)
                orphan = sorted({r for r in refs if r and r not in unit_ids})
                if orphan:
                    ds.issues.append(("WARNING", f"{len(orphan)} {col} serials are not listed "
                                                 f"in the Components sheet: {', '.join(orphan[:5])}"
                                                 + ("…" if len(orphan) > 5 else "")))
