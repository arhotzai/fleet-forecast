"""
Fleet ERP — web app
====================
Upload the master database workbook, optionally upload a fresh LLP report
(raw export), run the technical projection, and browse the results.
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
try:
    from streamlit_option_menu import option_menu
    _HAS_OPTION_MENU = True
except Exception:          # library not installed (e.g. offline corporate PC)
    option_menu = None
    _HAS_OPTION_MENU = False

sys.path.insert(0, str(Path(__file__).parent / "src"))
from fleet_erp.ingestion.loader import load          # noqa: E402
from fleet_erp.technical import engine               # noqa: E402
from fleet_erp.export.excel import build_workbook     # noqa: E402

import base64  # noqa: E402

# ── Palette (dark / Power BI) ─────────────────────────────────
BG0      = "#0B0F1A"   # page
PANEL    = "#141A27"   # cards / panels
PANEL2   = "#1B2231"
BORDER   = "#273046"
ACCENT   = "#2D7FF9"   # Aegean-ish blue
ACCENT2  = "#4FB2F0"
TEXT     = "#E6EAF2"
MUTED    = "#8A93A6"
C_LIMIT   = "#2D7FF9"
C_HARD    = "#17B3A3"
C_SOFT    = "#F2B134"
C_LLP     = "#E0557B"
PLOTLY_BG = "rgba(0,0,0,0)"

st.set_page_config(page_title="Fleet Forecast · AEGEAN", page_icon="✈️", layout="wide",
                   initial_sidebar_state="expanded")

_LOGO = Path(__file__).parent / "assets" / "aegean_logo.png"
_logo_b64 = base64.b64encode(_LOGO.read_bytes()).decode() if _LOGO.exists() else ""
_logo_html = (
    f'<img src="data:image/png;base64,{_logo_b64}" class="ff-logo" alt="AEGEAN"/>'
    if _logo_b64 else '<span class="ff-logo-txt">AEGEAN</span>'
)

st.markdown(f"""
<style>
  .stApp {{ background:
      radial-gradient(1200px 600px at 15% -10%, #16233f 0%, rgba(11,15,26,0) 55%),
      radial-gradient(1000px 500px at 95% -20%, #10263a 0%, rgba(11,15,26,0) 50%),
      {BG0}; }}
  header[data-testid="stHeader"] {{ background:rgba(0,0,0,0); }}
  .block-container {{ padding-top: 3rem; max-width: 1500px; }}
  html, body, [class*="css"] {{ font-family:'Inter','Segoe UI',system-ui,sans-serif; }}

  /* ── header ── */
  .ff-head {{
    display:flex; align-items:center; gap:2rem;
    background: linear-gradient(105deg, #0c1426 0%, #173f82 52%, #0c1426 100%);
    border:1px solid #2f4a7a; border-radius:20px;
    padding:2.1rem 2.4rem; margin-bottom:1.3rem; min-height:120px;
    box-shadow:0 14px 50px rgba(0,0,0,.5), inset 0 1px 0 rgba(255,255,255,.05);
    position:relative; overflow:hidden;
  }}
  .ff-head::after {{ content:""; position:absolute; left:0; right:0; bottom:0; height:4px;
    background:linear-gradient(90deg,{ACCENT} 0%, {ACCENT2} 45%, transparent 90%); }}
  .ff-logo {{ height:72px; filter:brightness(0) invert(1); opacity:.98; }}
  .ff-logo-txt {{ color:#fff; font-weight:900; font-size:2.2rem; letter-spacing:4px; }}
  .ff-divider {{ width:1px; height:76px; background:rgba(255,255,255,.18); }}
  .ff-title h1 {{ margin:0; font-size:3rem; font-weight:800; color:#fff; letter-spacing:.3px;
    line-height:1; }}
  .ff-title h1 .accent {{ color:{ACCENT2}; }}

  /* ── KPI tiles ── */
  .kpi-row {{ display:grid; grid-template-columns:repeat(5,1fr); gap:.8rem; margin:.2rem 0 1rem; }}
  .kpi {{ background:linear-gradient(180deg,{PANEL2} 0%,{PANEL} 100%);
    border:1px solid {BORDER}; border-radius:14px; padding:.9rem 1.05rem;
    position:relative; overflow:hidden; box-shadow:0 6px 20px rgba(0,0,0,.3); }}
  .kpi::before {{ content:""; position:absolute; left:0; top:0; bottom:0; width:3px;
    background:var(--c,{ACCENT}); }}
  .kpi .v {{ font-size:1.9rem; font-weight:800; color:{TEXT}; line-height:1.05; }}
  .kpi .l {{ font-size:.68rem; color:{MUTED}; text-transform:uppercase; letter-spacing:.7px;
    margin-top:.25rem; }}
  .kpi .s {{ font-size:.72rem; color:{ACCENT2}; margin-top:.15rem; }}

  /* ── tabs ── */
  .stTabs [data-baseweb="tab-list"] {{ gap:.4rem;
    background:linear-gradient(180deg,#141d30 0%, #0f1726 100%);
    border:1px solid {BORDER}; border-radius:16px; padding:.45rem;
    box-shadow:0 10px 30px rgba(0,0,0,.4), inset 0 1px 0 rgba(255,255,255,.04); }}
  .stTabs [data-baseweb="tab-highlight"], .stTabs [data-baseweb="tab-border"] {{ display:none; }}
  .stTabs [data-baseweb="tab"] {{ border-radius:11px; padding:.5rem 1.15rem; color:{MUTED};
    font-weight:600; font-size:.85rem; letter-spacing:.2px; border:1px solid transparent;
    transition:all .18s ease; }}
  .stTabs [data-baseweb="tab"]:hover {{ color:{TEXT};
    background:rgba(255,255,255,.04); border-color:rgba(255,255,255,.07);
    transform:translateY(-1px); }}
  .stTabs [aria-selected="true"] {{ color:#fff !important;
    background:linear-gradient(135deg,{ACCENT} 0%, #1f5fd0 100%) !important;
    border-color:rgba(255,255,255,.18) !important;
    box-shadow:0 6px 18px rgba(45,127,249,.5), inset 0 1px 0 rgba(255,255,255,.25);
    transform:translateY(-1px); }}
  .stTabs [aria-selected="true"]:hover {{ color:#fff !important;
    background:linear-gradient(135deg,{ACCENT2} 0%, {ACCENT} 100%) !important; }}

  /* ── panels / widgets ── */
  h2, h3, .stSubheader {{ color:{TEXT} !important; }}
  [data-testid="stDataFrame"] {{ border:1px solid {BORDER}; border-radius:12px; }}
  section[data-testid="stSidebar"] {{ background:#0c111c; border-right:1px solid {BORDER}; }}
  div[data-testid="stMetricValue"] {{ color:{TEXT}; }}
  .stButton>button[kind="primary"] {{ background:{ACCENT}; border:0; box-shadow:0 4px 14px rgba(45,127,249,.4); }}
  footer {{ visibility:hidden; }}
</style>
<div class="ff-head">
  {_logo_html}
  <div class="ff-divider"></div>
  <div class="ff-title"><h1>Fleet <span class="accent">Forecast</span></h1></div>
</div>
""", unsafe_allow_html=True)


def _secret(key, default=None):
    """Read a value from Streamlit secrets without erroring when none are set."""
    try:
        return st.secrets.get(key, default)
    except Exception:
        return default


# Optional confidential gate: set `app_password` in Streamlit secrets to require it.
_APP_PW = _secret("app_password")
if _APP_PW and not st.session_state.get("_authed"):
    _PLANE = ("M480 192H365.71L260.61 8.06A16.014 16.014 0 0 0 246.71 0h-65.5c-10.63 0-18.3 "
              "10.17-15.38 20.39L214.86 192H112l-43.2-57.6c-3.02-4.03-7.77-6.4-12.8-6.4H16.01"
              "C5.6 128-2.04 137.78.49 147.88L32 256 .49 364.12C-2.04 374.22 5.6 384 16.01 "
              "384H56c5.03 0 9.78-2.37 12.8-6.4L112 320h102.86l-51.03 171.61C160.91 501.83 "
              "168.58 512 179.21 512h65.5c5.74 0 11.04-3.08 13.9-8.06L365.71 320H480c35.35 0 "
              "96-28.65 96-64s-60.65-64-96-64z")
    st.markdown(f"""
    <style>
      .ff-hero {{ position:relative; margin:.2rem 0 1.1rem; border-radius:18px; overflow:hidden;
        background:linear-gradient(120deg,#0a1020 0%, #10274e 50%, #0a1020 100%);
        border:1px solid #2f4a7a; height:104px;
        box-shadow:0 14px 50px rgba(0,0,0,.5), inset 0 1px 0 rgba(255,255,255,.05); }}
      .ff-hero .sky {{ position:absolute; inset:0; width:100%; height:100%; }}
      .ff-arc {{ stroke-dasharray:5 9; animation:ffdash 16s linear infinite; }}
      @keyframes ffdash {{ to {{ stroke-dashoffset:-320; }} }}
      .ff-plane {{ animation:fffly 9s ease-in-out infinite; }}
      @keyframes fffly {{ 0%{{transform:translate(0,0)}} 50%{{transform:translate(11px,-6px)}}
        100%{{transform:translate(0,0)}} }}
      .ff-star {{ animation:fftw 4s ease-in-out infinite; }}
      @keyframes fftw {{ 0%,100%{{opacity:.25}} 50%{{opacity:.9}} }}
      /* login card */
      div[data-testid="stForm"] {{ background:linear-gradient(180deg,#18212f,#121825);
        border:1px solid #2b3550; border-radius:18px; padding:1.9rem 1.9rem 1.5rem;
        box-shadow:0 18px 50px rgba(0,0,0,.5); position:relative; overflow:hidden; }}
      div[data-testid="stForm"]::before {{ content:""; position:absolute; left:0; right:0; top:0;
        height:3px; background:linear-gradient(90deg,{ACCENT},{ACCENT2},transparent); }}
      div[data-testid="stForm"] [data-testid="stMarkdownContainer"] {{ text-align:center; }}
      div[data-testid="stForm"] [data-testid="stMarkdownContainer"] h4 {{ font-size:1.15rem; }}
      div[data-testid="stForm"] input {{ background:#0d1420 !important; border:1px solid #2b3550 !important;
        border-radius:11px !important; color:#E6EAF2 !important; text-align:center; }}
      .ff-confid {{ text-align:center; color:{MUTED}; font-size:.8rem; margin-top:.75rem; }}
    </style>
    <div class="ff-hero">
      <svg class="sky" viewBox="0 0 1200 100" preserveAspectRatio="xMidYMid slice">
        <defs>
          <radialGradient id="g1" cx="50%" cy="4%" r="82%">
            <stop offset="0%" stop-color="{ACCENT}" stop-opacity=".30"/>
            <stop offset="100%" stop-color="{ACCENT}" stop-opacity="0"/>
          </radialGradient>
        </defs>
        <rect width="1200" height="100" fill="url(#g1)"/>
        <circle class="ff-star" cx="300" cy="34" r="1.5" fill="#fff"/>
        <circle class="ff-star" cx="470" cy="62" r="1.8" fill="#fff" style="animation-delay:1s"/>
        <circle class="ff-star" cx="730" cy="30" r="1.4" fill="#fff" style="animation-delay:2s"/>
        <circle class="ff-star" cx="900" cy="60" r="1.7" fill="#fff" style="animation-delay:.5s"/>
        <circle class="ff-star" cx="1040" cy="34" r="1.4" fill="#fff" style="animation-delay:1.5s"/>
        <path class="ff-arc" d="M110 50 Q 600 26 1090 50" fill="none" stroke="{ACCENT2}"
              stroke-width="2.4" stroke-linecap="round" opacity=".9"/>
        <circle cx="110" cy="50" r="6" fill="#fff"/>
        <circle cx="1090" cy="50" r="6" fill="{ACCENT2}"/>
        <g class="ff-plane">
          <g transform="translate(580 20) rotate(2) scale(.072)">
            <path fill="#ffffff" d="{_PLANE}"/>
          </g>
        </g>
      </svg>
    </div>
    """, unsafe_allow_html=True)

    _lcol = st.columns([1, 1.25, 1])
    with _lcol[1]:
        with st.form("login"):
            st.markdown("#### 🔒 Enter password to continue")
            _pw = st.text_input("Password", type="password", label_visibility="collapsed",
                                placeholder="Enter password")
            if st.form_submit_button("Enter  →", type="primary", use_container_width=True):
                if _pw == str(_APP_PW):
                    st.session_state["_authed"] = True
                    st.rerun()
                else:
                    st.error("Incorrect password.")
        st.markdown("<div class='ff-confid'>🔒 Confidential — access restricted to "
                    "authorised Aegean personnel.</div>", unsafe_allow_html=True)
    st.stop()


@st.cache_data(show_spinner="Running projection…")
def _run(db_bytes: bytes, llp_bytes: bytes | None):
    ds = load(db_bytes, llp_bytes)
    if ds.errors:
        return ds, None
    proj = engine.run(ds)
    return ds, proj


# ── Navigation (lives in the left sidebar, under the upload) ──
NAV_ASK   = "Ask Claude"
NAV_DASH  = "SV Dashboard"
NAV_UTIL  = "Fleet & Utilisation Stats"
NAV_CAL   = "SV Calendar"
NAV_LOG   = "Event Log"
NAV_LLP   = "LLP Status"
NAV_AC    = "Aircraft Detail"
NAV_POWER = "Powerplant"
NAV_STATUS = "Current Fleet Status"
NAV_EDIT  = "Edit Data"
NAV_EXP   = "Export"
NAV_GUIDE = "Guide & Templates"
NAV_NOTES = "Data Notes"
NAV_ITEMS = [NAV_ASK, NAV_STATUS, NAV_UTIL, NAV_DASH, NAV_CAL, NAV_LOG, NAV_LLP, NAV_AC,
             NAV_POWER, NAV_EDIT, NAV_EXP, NAV_GUIDE, NAV_NOTES]
NAV_ICONS = ["chat-dots", "clipboard-data", "graph-up-arrow", "bar-chart-fill", "calendar3",
             "list-task", "wrench-adjustable", "airplane", "fan", "pencil-square",
             "box-arrow-down", "mortarboard", "info-circle"]

# professional sidebar navigation (streamlit-option-menu)
NAV_STYLES = {
    "container": {"padding": "4px 2px", "background-color": "transparent"},
    "menu-title": {"color": "#8A93A6", "font-size": "12px", "font-weight": "700",
                   "letter-spacing": "2.5px", "text-transform": "uppercase",
                   "text-align": "center", "margin-bottom": "10px"},
    "icon": {"color": "#8fa3c4", "font-size": "17px"},
    "nav-link": {"font-size": "15px", "font-weight": "600", "color": "#c3cee0",
                 "padding": "12px 16px", "margin": "5px 2px", "border-radius": "12px",
                 "border": "1px solid #232d42", "background-color": "#111a28",
                 "--hover-color": "#1a2740"},
    "nav-link-selected": {"background": "linear-gradient(135deg,#2D7FF9 0%,#1f5fd0 100%)",
                          "color": "#ffffff", "font-weight": "700", "border": "1px solid transparent",
                          "box-shadow": "0 6px 18px rgba(45,127,249,0.45)"},
}

# fallback nav styling (used when streamlit-option-menu isn't installed)
st.markdown("""<style>
  section[data-testid="stSidebar"] div[role="radiogroup"]{gap:.42rem;}
  section[data-testid="stSidebar"] div[role="radiogroup"] > label{
    display:flex;align-items:center;justify-content:center;width:100%;margin:0;
    padding:.75rem .9rem;border-radius:12px;cursor:pointer;background:#121a28;
    border:1px solid #232d42;transition:all .16s ease;}
  section[data-testid="stSidebar"] div[role="radiogroup"] > label:hover{
    background:#17223a;border-color:#3a5fa0;}
  section[data-testid="stSidebar"] div[role="radiogroup"] > label > div:first-child{display:none;}
  section[data-testid="stSidebar"] div[role="radiogroup"] > label p{
    font-size:1rem!important;font-weight:600;color:#aeb9cc;text-align:center;}
  section[data-testid="stSidebar"] div[role="radiogroup"] > label:has(input:checked){
    background:linear-gradient(135deg,#2D7FF9,#1f5fd0);border-color:transparent;}
  section[data-testid="stSidebar"] div[role="radiogroup"] > label:has(input:checked) p{color:#fff!important;}
  .nav-h{color:#8A93A6;font-size:.72rem;letter-spacing:2.5px;font-weight:700;
    text-transform:uppercase;margin:1rem 0 .55rem;text-align:center;}
</style>""", unsafe_allow_html=True)


# ── Sidebar: uploads + navigation ─────────────────────────────
with st.sidebar:
    st.header("Database")
    db_file = st.file_uploader("Database file (.xlsx) — any file name", type=["xlsx"], key="db",
                               help="Upload your Fleet Maintenance Status database. The file name "
                                    "doesn't matter — only its contents are read.")
    if db_file is None:
        st.caption("Upload your database file to begin — nothing is pre-loaded.")

    st.write("")
    if _HAS_OPTION_MENU:
        nav = option_menu(menu_title="Navigation", options=NAV_ITEMS, icons=NAV_ICONS,
                          menu_icon="grid-1x2-fill", default_index=0, styles=NAV_STYLES,
                          key="nav")
    else:
        st.markdown("<div class='nav-h'>Navigation</div>", unsafe_allow_html=True)
        nav = st.radio("Navigate", NAV_ITEMS, label_visibility="collapsed", key="nav")

    st.divider()
    run_clicked = st.button("↻ Re-run projection", type="primary", use_container_width=True)

# resolve the database source: in-app edits > uploaded file > bundled default.
# A new/changed upload supersedes (and clears) any in-app edits.
_uploaded_bytes = db_file.getvalue() if db_file is not None else None
if _uploaded_bytes is not None:
    _uh = hash(_uploaded_bytes)
    if st.session_state.get("_uploaded_hash") != _uh:
        st.session_state["_uploaded_hash"] = _uh
        st.session_state.pop("edited_db_bytes", None)
# No bundled/hardcoded data — the app is ALWAYS empty until a file is uploaded
# (or edited in-app). Nothing is pre-loaded.
if st.session_state.get("edited_db_bytes") is not None:
    db_bytes = st.session_state["edited_db_bytes"]
elif _uploaded_bytes is not None:
    db_bytes = _uploaded_bytes
else:
    db_bytes = None

if db_bytes is None:
    st.info("⬅ Upload your **Fleet Maintenance Status Database** (.xlsx) in the sidebar to begin.")
    st.stop()

# run automatically (cached by bytes); re-run button clears the cache
if run_clicked:
    _run.clear()
_dbhash = hash(db_bytes)
if run_clicked or st.session_state.get("_dbhash") != _dbhash:
    ds, proj = _run(db_bytes, None)
    st.session_state["ds"] = ds
    st.session_state["proj"] = proj
    st.session_state["_dbhash"] = _dbhash

ds = st.session_state.get("ds")
proj = st.session_state.get("proj")

# ── Validation feedback ───────────────────────────────────────
if ds.errors:
    st.error("Cannot run — please fix these and re-upload:")
    for e in ds.errors:
        st.write("• ", e)
    st.stop()
for w in ds.warnings:
    st.warning(w, icon="⚠️")

if proj is None or proj.events is None:
    st.stop()

events = proj.events
fleet = ds.get("Fleet")


def _active_today(_fl):
    """Aircraft in service today: Delivery ≤ today ≤ Redelivery (or Extension)."""
    if _fl is None or _fl.empty:
        return _fl
    d = _fl.copy()
    for _c in ("Delivery_Date", "Redelivery_Date", "Extension_Date"):
        if _c in d.columns:
            d[_c] = pd.to_datetime(d[_c], errors="coerce")
    _today = pd.Timestamp.today().normalize()
    _end = None
    if "Redelivery_Date" in d.columns:
        _end = d["Redelivery_Date"]
    if "Extension_Date" in d.columns:
        _end = d["Extension_Date"].fillna(_end) if _end is not None else d["Extension_Date"]
    m = pd.Series(True, index=d.index)
    if "Delivery_Date" in d.columns:
        m &= d["Delivery_Date"].isna() | (d["Delivery_Date"] <= _today)
    if _end is not None:
        m &= _end.isna() | (_end >= _today)
    return d[m]


fleet_active = _active_today(fleet)

# ── KPI tiles (Power BI style) ────────────────────────────────
_n_ac = fleet_active["Tail"].nunique() if not fleet_active.empty else 0
_n_units = proj.unit_status["Unit_ID"].nunique() if (proj.unit_status is not None and not proj.unit_status.empty) else 0
next12 = events[pd.to_datetime(events["Date"]) <= proj.dates[0] + pd.offsets.MonthEnd(12)] if not events.empty else events
_n_llp = int((events["Trigger"] == "LLP_REPLACEMENT").sum()) if not events.empty else 0


def _tile(value, label, sub="", color=ACCENT):
    return (f'<div class="kpi" style="--c:{color}"><div class="v">{value}</div>'
            f'<div class="l">{label}</div><div class="s">{sub}</div></div>')


def _tile_break(label, items, color=ACCENT, total=None):
    """A KPI card whose body is a compact 'label … count' breakdown list."""
    head = (f'<div class="v" style="font-size:1.7rem">{total}</div>'
            if total is not None else "")
    rows = "".join(
        f'<div style="display:flex;justify-content:space-between;gap:1rem;font-size:.82rem;'
        f'color:{TEXT};line-height:1.7"><span style="color:#aab4c6">{k}</span>'
        f'<b>{v}</b></div>' for k, v in items)
    return (f'<div class="kpi" style="--c:{color}">{head}'
            f'<div class="l" style="margin-bottom:.35rem">{label}</div>{rows}</div>')


def _blank(df: pd.DataFrame) -> pd.DataFrame:
    """Show empty cells instead of None / NaN / NaT in a table (any text dtype)."""
    d = df.copy()
    for c in d.columns:
        col = d[c]
        if pd.api.types.is_datetime64_any_dtype(col):
            d[c] = col.dt.strftime("%Y-%m-%d").where(col.notna(), "")
        elif not pd.api.types.is_numeric_dtype(col):   # object, string, category…
            d[c] = col.astype(object).where(col.notna(), "")
    return d


def _cols(df: pd.DataFrame, wanted: list[str]) -> pd.DataFrame:
    """Select only the columns that exist (resilient to schema changes)."""
    return df[[c for c in wanted if c in df.columns]]


def _fmt(v, as_int=False):
    if v is None:
        return ""
    try:
        if pd.isna(v):          # catches NaN and NaT scalars
            return ""
    except (TypeError, ValueError):
        pass
    if hasattr(v, "strftime"):
        return v.strftime("%Y-%m-%d")
    if as_int:
        try:
            return f"{int(round(float(v))):,}"
        except (TypeError, ValueError):
            return str(v)
    return str(v)


def _llp_table_html(df: pd.DataFrame, cap: int = 500) -> tuple[str, int]:
    """Render LLPs as a dark table with a soft two-tone life bar (consumed ▸ to-go)."""
    CONS = "#6a9bd8"   # light blue — consumed
    LEFT = "#5bbf86"   # soft green — remaining (to-go)
    rows = []
    for _, r in df.head(cap).iterrows():
        lim, togo, slsv = r.get("Limit_Cycles"), r.get("ToGo"), r.get("SLSV")
        try:
            tg = max(0.0, min(100.0, float(togo) / float(lim) * 100)) if lim else 0.0
        except (TypeError, ValueError, ZeroDivisionError):
            tg = 0.0
        cons = 100.0 - tg
        bar = (
            '<div style="display:flex;align-items:center;gap:8px">'
            '<div style="flex:0 0 140px;height:13px;border-radius:4px;overflow:hidden;'
            'background:#111725;display:flex">'
            f'<div style="width:{cons:.1f}%;background:{CONS}"></div>'
            f'<div style="width:{tg:.1f}%;background:{LEFT}"></div></div>'
            f'<span style="color:{LEFT};font-size:.74rem">{tg:.0f}% left</span></div>'
        )
        rows.append(
            f"<tr><td>{_fmt(r.get('Unit_ID'))}</td><td>{_fmt(r.get('LLP_PN'))}</td>"
            f"<td>{_fmt(r.get('LLP_SN'))}</td><td>{_fmt(r.get('Description'))}</td>"
            f"<td class='num'>{_fmt(lim, True)}</td><td class='num'>{_fmt(slsv, True)}</td>"
            f"<td class='num'>{_fmt(togo, True)}</td><td style='min-width:230px'>{bar}</td>"
            f"<td>{_fmt(r.get('Projected_SV'))}</td><td>{_reason_badge(r.get('Reason'))}</td></tr>"
        )
    head = ("<tr><th>Unit_ID</th><th>LLP PN</th><th>LLP SN</th><th>Description</th>"
            "<th class='num'>Limit</th><th class='num'>SLSV</th><th class='num'>To-go</th>"
            "<th>Life&nbsp;(consumed&nbsp;▸&nbsp;left)</th><th>Projected SV</th><th>Reason</th></tr>")
    style = ("<style>.llp-tbl{width:100%;border-collapse:collapse;font-size:.82rem;color:#E6EAF2;white-space:nowrap}"
             ".llp-tbl th{text-align:left;color:#8A93A6;font-weight:600;padding:8px 12px;"
             "border-bottom:1px solid #273046;position:sticky;top:0;background:#0f1520;z-index:1}"
             ".llp-tbl td{padding:6px 12px;border-bottom:1px solid #1b2231}"
             ".llp-tbl td.num,.llp-tbl th.num{text-align:right}</style>")
    html = (style + '<div style="max-height:520px;overflow:auto;border:1px solid #273046;'
            f'border-radius:12px">\n<table class="llp-tbl">{head}{"".join(rows)}</table></div>')
    return html, min(len(df), cap)


def _reason_badge(val) -> str:
    """Vibrant colour-coded pill for a projection 'Reason' value."""
    s = _fmt(val)
    if not s:
        return ""
    low = s.lower()
    c = (C_HARD if "hard link" in low else
         C_SOFT if "soft link" in low else
         C_LLP if ("llp" in low or "replace" in low) else
         C_LIMIT)
    return (f'<span style="display:inline-block;padding:3px 13px;border-radius:999px;'
            f'background:{c}22;color:{c};border:1px solid {c}66;font-size:.74rem;'
            f'font-weight:700;white-space:nowrap;box-shadow:0 0 0 1px {c}22">{s}</span>')


def _signed_pill(v) -> str:
    """Green (life remaining) / red (overdue) pill for a signed numeric value."""
    try:
        f = float(v)
    except (TypeError, ValueError):
        return _fmt(v)
    if pd.isna(f):
        return ""
    pos = f >= 0
    c = "#2ecc71" if pos else "#ff6b6b"
    txt = (f"{int(f):+,}" if f == int(f) else f"{f:+,.1f}")
    return (f'<span style="display:inline-flex;align-items:center;gap:6px;padding:3px 12px;'
            f'border-radius:999px;background:{c}1e;color:{c};border:1px solid {c}55;'
            f'font-weight:700;font-size:.78rem;white-space:nowrap">'
            f'<span style="width:7px;height:7px;border-radius:50%;background:{c};'
            f'box-shadow:0 0 6px {c}">&nbsp;</span>{txt}</span>')


def _pretty_table(df: pd.DataFrame, columns: list[str], reason_col: str = "Reason",
                  height: int = 360, cap: int = 200, signed_cols: list[str] | None = None,
                  labels: dict[str, str] | None = None) -> str:
    """Render a dark, professional HTML table: sticky header, zebra rows, hover,
    tabular numbers, colour-coded 'Reason' pills, and green/red signed pills."""
    signed = set(signed_cols or [])
    labels = labels or {}
    cols = [c for c in columns if c in df.columns]
    if not cols:
        return ""
    d = df[cols].head(cap)
    numeric = {c for c in cols if pd.api.types.is_numeric_dtype(d[c])}

    def _lab(c):
        return labels.get(c, str(c).replace("_", " "))

    def _numfmt(v):
        try:
            f = float(v)
        except (TypeError, ValueError):
            return _fmt(v)
        if pd.isna(f):
            return ""
        return f"{int(f):,}" if f == int(f) else f"{f:,.1f}"

    def _cell(c, v):
        if c == reason_col:
            return f"<td>{_reason_badge(v)}</td>"
        if c in signed:
            return f"<td class='num'>{_signed_pill(v)}</td>"
        if c in numeric:
            return f"<td class='num'>{_numfmt(v)}</td>"
        return f"<td>{_fmt(v)}</td>"

    head = "".join((f"<th class='num'>{_lab(c)}</th>" if c in numeric else f"<th>{_lab(c)}</th>")
                   for c in cols)
    body = "".join("<tr>" + "".join(_cell(c, r[c]) for c in cols) + "</tr>"
                   for _, r in d.iterrows())
    style = ("<style>"
             ".ff-tbl{width:100%;border-collapse:separate;border-spacing:0;font-size:.82rem;"
             "color:#E6EAF2;white-space:nowrap}"
             ".ff-tbl th{text-align:left;color:#9aa6bd;font-weight:700;padding:11px 16px;"
             "text-transform:uppercase;letter-spacing:.5px;font-size:.68rem;"
             "background:linear-gradient(180deg,#141d30,#0f1726);"
             "border-bottom:1px solid #2b3550;position:sticky;top:0;z-index:1}"
             ".ff-tbl td{padding:9px 16px;border-bottom:1px solid #1a2130}"
             ".ff-tbl tbody tr:nth-child(even) td{background:rgba(255,255,255,.016)}"
             ".ff-tbl tbody tr:hover td{background:rgba(45,127,249,.10)}"
             ".ff-tbl td.num,.ff-tbl th.num{text-align:right;font-variant-numeric:tabular-nums}"
             ".ff-tbl tbody tr:first-child td{border-top:0}"
             "</style>")
    return (style + f'<div style="max-height:{height}px;overflow:auto;border:1px solid #273046;'
            f'border-radius:14px;box-shadow:0 8px 28px rgba(0,0,0,.35)">\n'
            f'<table class="ff-tbl"><thead><tr>{head}</tr></thead>'
            f'<tbody>{body}</tbody></table></div>')


# ── Global fleet cards: clean counts by big category (Total / NEO / CEO / ATR) ──
_NEO_C, _CEO_C, _ATR_C = "#16489C", "#2D7FF9", "#3FB06B"


def _big_cat(name):
    f = str(name).upper()
    if "NEO" in f:
        return "NEO"
    if "CEO" in f:
        return "CEO"
    if "ATR" in f or "DASH" in f or "DH" in f or "Q400" in f or "D100" in f:
        return "ATR"
    return "Other"


# aircraft by big category (active today)
_ac_cat = {"NEO": 0, "CEO": 0, "ATR": 0}
if fleet_active is not None and not fleet_active.empty:
    for _fam, _cnt in fleet_active["Aircraft_Family"].value_counts().items():
        _c = _big_cat(_fam)
        if _c in _ac_cat:
            _ac_cat[_c] += int(_cnt)

# spare engines by big category (engine model → category); untyped excluded
_ENG_CAT = {"PW1100": "NEO", "PW1500": "NEO", "LEAP": "NEO",
            "V2500": "CEO", "CFM56": "CEO", "PW127": "ATR"}
_sp_cat = {"NEO": 0, "CEO": 0, "ATR": 0}
_n_spare_apu = 0
_comp_g = ds.get("Components")
if _comp_g is not None and not _comp_g.empty and "Installed_On_Tail" in _comp_g.columns:
    _free = (_comp_g["Installed_On_Tail"].isna()
             | _comp_g["Installed_On_Tail"].astype(str).str.strip().isin(["", "nan", "None"]))
    _eng_sp = _comp_g[(_comp_g["Unit_Kind"] == "ENGINE") & _free]
    _n_spare_apu = int(((_comp_g["Unit_Kind"] == "APU") & _free).sum())
    for _et, _cnt in _eng_sp["Unit_Type"].value_counts().items():
        _c = _ENG_CAT.get(str(_et))
        if _c in _sp_cat:
            _sp_cat[_c] += int(_cnt)

# Data reference / snapshot date — the starting point of the whole projection
_cfg_asof = ds.get("Config")
_asof_str = ""
if _cfg_asof is not None and not _cfg_asof.empty and "Setting" in _cfg_asof.columns:
    _m = _cfg_asof[_cfg_asof["Setting"].astype(str) == "As_Of_Date"]
    if not _m.empty:
        _ad = pd.to_datetime(_m["Value"].iloc[0], errors="coerce")
        _asof_str = _ad.strftime("%d %b %Y") if pd.notna(_ad) else str(_m["Value"].iloc[0])
if _asof_str:
    st.markdown(
        f'<div style="display:inline-block;background:#111a28;border:1px solid #2b3550;'
        f'border-left:3px solid {ACCENT};border-radius:10px;padding:.45rem .9rem;margin:.1rem 0 .7rem;'
        f'font-size:.85rem;color:{TEXT};">📅 <b>Data reference (snapshot): {_asof_str}</b>'
        f'<span style="color:{MUTED};">  — the projection\'s starting point. '
        f'All current status (TSN/CSN/SLSV) is as of this date.</span></div>',
        unsafe_allow_html=True)

# "live" pulsing green dot for the active-fleet card
st.markdown("""<style>
  .live-dot{display:inline-block;width:9px;height:9px;border-radius:50%;background:#2ecc71;
    margin-right:6px;vertical-align:middle;box-shadow:0 0 0 0 rgba(46,204,113,.7);
    animation:livepulse 1.6s infinite;}
  @keyframes livepulse{0%{box-shadow:0 0 0 0 rgba(46,204,113,.6)}
    70%{box-shadow:0 0 0 8px rgba(46,204,113,0)}100%{box-shadow:0 0 0 0 rgba(46,204,113,0)}}
</style>""", unsafe_allow_html=True)

# Row 1 — aircraft: LIVE total fleet (active today) + NEO / CEO / ATR
_LIVE_GREEN = "#2ecc71"
st.markdown(
    '<div style="display:grid;grid-template-columns:repeat(4,1fr);gap:.8rem;margin:.2rem 0 .6rem;">'
    + _tile(f"{_n_ac}", '<span class="live-dot"></span>Live total fleet', "active today",
            _LIVE_GREEN)
    + _tile(f"{_ac_cat['NEO']}", "NEO", "aircraft", _NEO_C)
    + _tile(f"{_ac_cat['CEO']}", "CEO", "aircraft", _CEO_C)
    + _tile(f"{_ac_cat['ATR']}", "ATR", "aircraft", _ATR_C)
    + '</div>', unsafe_allow_html=True)

# Row 2 — spare engines by category + spare APUs
st.markdown(
    '<div style="display:grid;grid-template-columns:repeat(4,1fr);gap:.8rem;margin:0 0 1rem;">'
    + _tile(f"{_sp_cat['NEO']}", "NEO spare engines", "PW1100", _NEO_C)
    + _tile(f"{_sp_cat['CEO']}", "CEO spare engines", "V2500", _CEO_C)
    + _tile(f"{_sp_cat['ATR']}", "ATR spare engines", "PW127", _ATR_C)
    + _tile(f"{_n_spare_apu}", "Spare APUs", "all types", C_SOFT)
    + '</div>', unsafe_allow_html=True)

# ── SV Calendar (SV_PLAN_v1) ──────────────────────────────────
if nav == NAV_CAL:
    st.subheader("Shop-visit calendar — month × aircraft")
    v1 = proj.sv_plan_v1
    if v1 is None or v1.empty:
        st.info("No calendar available.")
    else:
        d = v1.copy()
        d["DATE"] = pd.to_datetime(d["DATE"]).dt.strftime("%b %Y")
        ac_cols = [c for c in d.columns if c != "DATE"]
        pick = st.multiselect("Aircraft", ac_cols, [], key="v1ac")
        show_cols = ["DATE"] + (pick if pick else ac_cols)
        only_active = st.checkbox("Only months with shop visits", value=True, key="v1active")
        view = d[show_cols]
        if only_active:
            view = view[view[[c for c in show_cols if c != "DATE"]].notna().any(axis=1)]
        st.caption(f"{len(view)} months shown. Each cell lists the shop visits due that month.")
        st.dataframe(_blank(view), use_container_width=True, hide_index=True, height=500)

# ── Event Log ─────────────────────────────────────────────────
if nav == NAV_LOG:
    st.subheader("Projected shop-visit / replacement events")
    if events.empty:
        st.info("No events projected.")
    else:
        colf = st.columns(4)
        trig = colf[0].multiselect("Trigger", sorted(events["Trigger"].unique()), [])
        kind = colf[1].multiselect("Unit kind", sorted(events["Unit_Kind"].unique()), [])
        tails = colf[2].multiselect("Tail", sorted(events["Tail"].unique()), [])
        horizon_mo = colf[3].slider("Within months", 6, len(proj.dates), min(60, len(proj.dates)))

        f = events.copy()
        f = f[pd.to_datetime(f["Date"]) <= proj.dates[0] + pd.offsets.MonthEnd(horizon_mo)]
        if trig:
            f = f[f["Trigger"].isin(trig)]
        if kind:
            f = f[f["Unit_Kind"].isin(kind)]
        if tails:
            f = f[f["Tail"].isin(tails)]
        disp = f.copy()
        disp["Date"] = pd.to_datetime(disp["Date"]).dt.strftime("%d %b %Y")
        note = f"{len(f):,} events"
        if len(f) > 200:
            note += "  ·  previewing the first 200 — use the CSV below for the full list"
        st.caption(note)
        st.markdown(_pretty_table(disp, ["Date", "Tail", "Position", "Unit_ID", "Unit_Type",
                                         "Event", "Driver", "SLSV_at_SV", "Reason"], height=520),
                    unsafe_allow_html=True)
        st.write("")
        st.download_button("⬇ Download event log (CSV)", events.to_csv(index=False).encode(),
                           "event_log.csv", "text/csv")

# ── SV Dashboard ──────────────────────────────────────────────
_TRIG_COLORS = {"LIMIT": C_LIMIT, "HARD_LINK": C_HARD, "SOFT_BUNDLE": C_SOFT,
                "LLP_REPLACEMENT": C_LLP}
# friendly category labels for the legend
TRIG_LABEL = {"LIMIT": "Reached limit", "HARD_LINK": "Hard link",
              "SOFT_BUNDLE": "Soft link", "LLP_REPLACEMENT": "LLP replacement"}
_TRIG_COLORS_F = {"Reached limit": C_LIMIT, "Hard link": C_HARD,
                  "Soft link": C_SOFT, "LLP replacement": C_LLP}

if nav == NAV_DASH:
    st.subheader("SV Dashboard — upcoming events by type")
    if events.empty:
        st.info("No events.")
    else:
        ev = events.copy()
        ev["Date"] = pd.to_datetime(ev["Date"])
        kind_label = {"AIRFRAME": "Airframe", "ENGINE": "Engine", "APU": "APU",
                      "PROPELLER": "Propeller"}
        import re as _re
        _LG = _re.compile(r"\b(MLG|NOSE\s*LG|LG|LANDING)\b", _re.I)

        def _scope(trig, kind, event):
            if trig == "LLP_REPLACEMENT":
                return "LLP"
            if kind == "AIRFRAME" and _LG.search(str(event)):
                return "Landing gear"
            return kind_label.get(kind, kind)

        ev["Scope"] = [_scope(t, k, e) for t, k, e in
                       zip(ev["Trigger"], ev["Unit_Kind"], ev["Event"])]
        # map each event to its aircraft family (via the tail it flies on)
        _fammap = fleet.set_index("Tail")["Aircraft_Family"].to_dict() if not fleet.empty else {}
        ev["Aircraft_Type"] = ev["Tail"].map(_fammap)

        f1, f2, f3, f4 = st.columns([1.2, 1.2, 1.8, .9])
        scope_sel = f1.multiselect("Category",
                                   ["Airframe", "Landing gear", "Engine", "APU", "Propeller", "LLP"],
                                   default=["Airframe"], key="sv_scope_ms",
                                   placeholder="All categories")
        _families = sorted(ev["Aircraft_Type"].dropna().unique().tolist())
        ac_sel = f2.multiselect("Aircraft type", _families, key="sv_actype_ms",
                                placeholder="All types")

        pool = ev if not scope_sel else ev[ev["Scope"].isin(scope_sel)]
        if ac_sel:
            pool = pool[pool["Aircraft_Type"].isin(ac_sel)]
        counts = pool["Event"].value_counts()
        sv_sel = f3.multiselect("SV type", list(counts.index), key="sv_type_ms",
                                format_func=lambda x: f"{x}   ({counts[x]})",
                                placeholder="All SV types in selection")
        horizon_y = f4.slider("Within next (years)", 1, 25, 10, key="sv_horizon")

        as_of = proj.dates[0]
        cutoff = as_of + pd.offsets.MonthEnd(horizon_y * 12)
        sel = pool if not sv_sel else pool[pool["Event"].isin(sv_sel)]
        sel = sel[(sel["Date"] >= as_of) & (sel["Date"] <= cutoff)].copy()

        total = len(sel)
        if total:
            nxt = sel["Date"].min()
            nxt_tail = sel.sort_values("Date")["Tail"].iloc[0]
            peak = int(sel["Date"].dt.year.value_counts().idxmax())
        else:
            nxt, nxt_tail, peak = None, "—", "—"

        st.markdown(
            '<div class="kpi-row">'
            + _tile(f"{total:,}", "Upcoming events", f"next {horizon_y} years", ACCENT)
            + _tile(nxt.strftime("%b %Y") if nxt is not None else "—", "Next due",
                    f"on {nxt_tail}", ACCENT2)
            + _tile(f"{sel['Tail'].nunique()}", "Aircraft affected", "", C_HARD)
            + _tile(str(peak), "Peak year", "most events", C_SOFT)
            + '</div>', unsafe_allow_html=True)

        if total:
            title_what = (", ".join(sv_sel) if sv_sel else
                          ", ".join(scope_sel) if scope_sel else "all")
            # chart title lives OUTSIDE the figure so it never collides with the legend
            st.markdown(
                f'<div style="font-weight:700;color:{TEXT};font-size:1.02rem;'
                f'margin:.3rem 0 -.2rem;">Upcoming “{title_what}” shop visits — timeline by aircraft</div>',
                unsafe_allow_html=True)

            order = sel.groupby("Tail")["Date"].min().sort_values().index.tolist()
            tl = sel.copy()
            tl["When"] = tl["Date"].dt.strftime("%b %Y")
            tl["Category"] = tl["Trigger"].map(TRIG_LABEL).fillna(tl["Trigger"])
            figc = px.scatter(tl, x="Date", y="Tail", color="Category",
                              color_discrete_map=_TRIG_COLORS_F,
                              category_orders={"Tail": order},
                              hover_data={"Event": True, "Reason": True, "When": True,
                                          "Date": False, "Tail": False, "Category": False})
            figc.update_traces(marker=dict(size=12, line=dict(width=1, color=BG0)), opacity=.95)
            figc.update_layout(template="plotly_dark",
                               height=min(900, max(360, len(order) * 24 + 130)),
                               plot_bgcolor=PLOTLY_BG, paper_bgcolor=PLOTLY_BG, font=dict(color=TEXT),
                               legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0, title=""),
                               margin=dict(l=10, r=10, t=36, b=10), xaxis_title="")
            figc.update_xaxes(gridcolor=BORDER, zeroline=False)
            figc.update_yaxes(gridcolor=BORDER, zeroline=False, title="")
            st.plotly_chart(figc, use_container_width=True)

            st.markdown("**Upcoming occurrences**")
            disp = sel.sort_values("Date")
            disp = disp.assign(Due=disp["Date"].dt.strftime("%d %b %Y"))
            st.markdown(_pretty_table(disp, ["Due", "Tail", "Position", "Unit_ID", "Event", "Reason"],
                                      height=360), unsafe_allow_html=True)
            st.write("")
            st.download_button("⬇ Download these events (CSV)", sel.to_csv(index=False).encode(),
                               "sv_upcoming.csv", "text/csv")
        else:
            st.info("No events of this type within the chosen horizon.")

# ── LLP Status ────────────────────────────────────────────────
if nav == NAV_LLP:
    st.subheader("Life-limited parts — current status")
    llp = ds.get("LLP_Registry")
    if llp.empty:
        st.info("No LLP data in the database.")
    else:
        data = llp.copy()
        # enrich with the PROJECTED first replacement (date + reason) from the engine
        le = events[events["Event_ID"].astype(str).str.startswith("LLP_")].copy()
        if not le.empty:
            le["LLP_SN"] = le["Event_ID"].astype(str).str[4:]
            le["Date"] = pd.to_datetime(le["Date"])
            firstsv = (le.sort_values("Date").groupby(["Unit_ID", "LLP_SN"])
                       .agg(Projected_SV=("Date", "first"), Reason=("Reason", "first")).reset_index())
            data = data.merge(firstsv, on=["Unit_ID", "LLP_SN"], how="left")
        else:
            data["Projected_SV"] = pd.NaT
            data["Reason"] = None

        for c in ("Limit_Cycles", "SLSV", "ToGo"):
            data[c] = pd.to_numeric(data[c], errors="coerce")
        data["To-go %"] = (data["ToGo"] / data["Limit_Cycles"] * 100).round(1)
        data["Life consumed %"] = (100 - data["To-go %"]).round(1)

        serials = sorted(data["Unit_ID"].dropna().unique())
        pick = st.multiselect("Engine / APU serial", serials, [], key="llp_serial",
                              help="Pick one (or a few) to see the per-part life bars.")
        view = (data[data["Unit_ID"].isin(pick)] if pick else data).sort_values("ToGo")

        cap = len(view) if pick else 200
        html, shown = _llp_table_html(view, cap=cap)
        note = f"{len(view):,} LLPs, nearest to limit first."
        if shown < len(view):
            note += f"  Showing the {shown} closest to limit — use the CSV for the full list."
        note += "  Projected SV + Reason come from the projection; Report ETA is in the CSV."
        st.caption(note)
        st.markdown(html, unsafe_allow_html=True)
        st.download_button("⬇ Download LLP status (CSV)", view.to_csv(index=False).encode(),
                           "llp_status.csv", "text/csv")

# ── Aircraft Detail ───────────────────────────────────────────
if nav == NAV_AC:
    st.subheader("Per-aircraft view")
    if fleet.empty:
        st.info("No fleet.")
    else:
        tail = st.selectbox("Aircraft", sorted(fleet["Tail"].dropna().unique()))
        row = fleet[fleet["Tail"] == tail].iloc[0]
        m = st.columns(5)
        m[0].metric("Family", str(row.get("Aircraft_Family", "—")))
        m[1].metric("LH ENG", str(row.get("LHENG_ESN", "—")))
        m[2].metric("RH ENG", str(row.get("RHENG_ESN", "—")))
        m[3].metric("APU", str(row.get("APU_SN", "—")))
        rd = pd.to_datetime(row.get("Redelivery_Date"), errors="coerce")
        m[4].metric("Redelivery", rd.strftime("%b %Y") if pd.notna(rd) else "—")
        view = st.radio("View", ["Shop visits (events)", "Month-by-month sheet"],
                        horizontal=True, key="ac_view")

        if view == "Shop visits (events)":
            sub = events[events["Tail"] == tail].copy()
            if sub.empty:
                st.info("No projected events for this aircraft.")
            else:
                sub["Date"] = pd.to_datetime(sub["Date"]).dt.strftime("%d %b %Y")
                st.caption(f"{len(sub)} projected events for {tail}.")
                st.markdown(_pretty_table(sub, ["Date", "Position", "Unit_ID", "Event",
                                                "Driver", "SLSV_at_SV", "Reason"], height=460),
                            unsafe_allow_html=True)
                st.write("")
                st.download_button("⬇ Events (CSV)", sub.to_csv(index=False).encode(),
                                   f"{tail}_events.csv", "text/csv")
        else:
            frame = (proj.aircraft_frames or {}).get(tail)
            if frame is None or frame.empty:
                st.info("No monthly data for this aircraft.")
            else:
                st.caption("Full projection like the Excel per-aircraft sheet: utilisation + every "
                           "counter, month by month. A counter showing **0** marks a shop visit that "
                           "month (it just reset). The downloadable Excel highlights those months in red.")
                disp = frame.copy()
                disp["DATE"] = pd.to_datetime(disp["DATE"]).dt.strftime("%b %Y")
                # optional: show only months where some event happens
                if st.checkbox("Only show shop-visit months", key="ac_sv_only"):
                    slsv_cols = [c for c in disp.columns if "SLSV" in str(c)]
                    num = disp[slsv_cols].apply(pd.to_numeric, errors="coerce")
                    disp = disp[(num == 0).any(axis=1)]
                st.dataframe(disp, use_container_width=True, hide_index=True, height=560)
                st.download_button("⬇ Download this aircraft sheet (CSV)",
                                   frame.to_csv(index=False).encode(), f"{tail}.csv", "text/csv")

# ── Ask Claude (home) ─────────────────────────────────────────
if nav == NAV_ASK:
    st.markdown(
        f"""<div style="padding:.4rem 0 1rem">
          <div style="font-size:1.7rem;font-weight:800;color:{TEXT};">
            ✨ Your fleet, on demand — <span style="color:{ACCENT2}">what can I bring you?</span>
          </div>
          <div style="color:{MUTED};margin-top:.3rem;font-size:.95rem;">
            Ask in plain language and I'll pull it straight from the projection — numbers, tables, rankings, timelines.
          </div>
        </div>""", unsafe_allow_html=True)
    st.caption("Try: “which 5 aircraft have the most shop visits in the next 3 years?” · "
               "“list engine LLP replacements due before 2030 by tail” · "
               "“how many APU shop visits per year?” · “what's due on SX-NEA in the next 18 months?”")
    import os
    try:
        from fleet_erp.ui import assistant as _assistant
        import anthropic
        _ASK_OK = True
    except Exception:
        _ASK_OK = False

    if not _ASK_OK:
        st.info("💬 **Ask Claude** needs the `anthropic` library, which isn't installed in this "
                "environment. Everything else in the app works normally — this tab is just hidden "
                "here. (It's available on the online version.)")
        st.stop()

    api_key = (os.environ.get("ANTHROPIC_API_KEY") or _secret("ANTHROPIC_API_KEY")
               or st.session_state.get("anthropic_key"))
    if not api_key:
        entered = st.text_input("Anthropic API key", type="password",
                                help="Or set ANTHROPIC_API_KEY before launching the app.")
        if entered:
            st.session_state["anthropic_key"] = entered
            api_key = entered

    if not api_key:
        st.info("Enter an Anthropic API key (or set the ANTHROPIC_API_KEY environment variable) to use the assistant.")
    else:
        st.session_state.setdefault("chat_display", [])
        st.session_state.setdefault("chat_hist", [])

        # input box lives up top (not docked at the bottom of the page)
        with st.form("ask_form", clear_on_submit=True):
            ci, cb = st.columns([7, 1])
            q = ci.text_input("Ask", placeholder="Ask about the projection…",
                              label_visibility="collapsed")
            ask = cb.form_submit_button("Ask ↑", type="primary", use_container_width=True)
        if st.session_state["chat_display"] and st.button("Clear chat"):
            st.session_state["chat_display"] = []
            st.session_state["chat_hist"] = []
            st.rerun()

        if ask and q:
            st.session_state["chat_display"].append({"role": "user", "text": q, "df": None})
            st.session_state["chat_hist"].append({"role": "user", "content": q})
            with st.spinner("Claude is working…"):
                try:
                    # Accept-Encoding: identity avoids a response-decompression bug
                    # in some local HTTP-library builds.
                    client = anthropic.Anthropic(api_key=api_key,
                                                 default_headers={"Accept-Encoding": "identity"})
                    frames = _assistant.available_frames(ds, proj)
                    text, df, _ = _assistant.ask(client, frames, st.session_state["chat_hist"])
                except anthropic.AuthenticationError:
                    text, df = "Authentication failed — check your API key.", None
                except Exception as exc:
                    text, df = f"Error: {exc}", None
            st.session_state["chat_display"].append({"role": "assistant", "text": text or "(no answer)", "df": df})
            st.rerun()

        # conversation shows directly under the input (newest last)
        for m in st.session_state["chat_display"]:
            with st.chat_message(m["role"]):
                st.markdown(m["text"])
                if m.get("df") is not None and not m["df"].empty:
                    st.dataframe(m["df"], use_container_width=True, hide_index=True, height=300)
                    st.download_button("⬇ CSV", m["df"].to_csv(index=False).encode(),
                                       "answer.csv", "text/csv", key=f"dl_{id(m)}")

# ── Utilisation & Statistics ──────────────────────────────────
if nav == NAV_UTIL:
    st.subheader("Utilisation assumptions & fleet statistics")
    util = ds.get("Utilization")
    if util.empty:
        st.info("No utilisation data in the database.")
    else:
        u = util.copy()
        for c in ["Annual_FH", "Annual_FC", "FH_FC_Ratio", "Annual_Evolution_Pct",
                  "Min_Annual_FH", "Max_Annual_FH", "AOG_Lead_In_Months"]:
            if c in u.columns:
                u[c] = pd.to_numeric(u[c], errors="coerce")
        _MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
                   "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
        air = u[u["Unit_Kind"] == "AIRFRAME"]

        def _sty(fig, title="", h=320, legend=False):
            fig.update_layout(template="plotly_dark", height=h, plot_bgcolor=PLOTLY_BG,
                              paper_bgcolor=PLOTLY_BG, font=dict(color=TEXT), title=title,
                              margin=dict(l=10, r=10, t=46, b=10), showlegend=legend,
                              legend=dict(title="", orientation="v", x=1.01, y=1,
                                          bgcolor="rgba(0,0,0,0)"))
            fig.update_xaxes(gridcolor=BORDER, zeroline=False)
            fig.update_yaxes(gridcolor=BORDER, zeroline=False)
            return fig

        # colour each aircraft type by family group (distinct shades within a group)
        _GREEN, _BLUE, _DEEP = "#3FB06B", "#2D7FF9", "#16489C"
        _FAM_COLORS = {
            "A320ceo": "#2D7FF9", "A321ceo": "#5AA7FF", "A320ceoSharklets": "#1C6FE0",
            "A320neo": "#16489C", "A321neo": "#2E63C8",
            "ATR72": "#3FB06B", "ATR42": "#2E8B57", "ATR72XT": "#79D6A4",
        }

        def _fam_color(fam):
            if fam in _FAM_COLORS:
                return _FAM_COLORS[fam]
            f = str(fam).upper()
            if "ATR" in f or "D100" in f or "DASH" in f or "DH" in f or "Q400" in f:
                return _GREEN
            if "NEO" in f:
                return _DEEP
            if "CEO" in f or f.startswith("A3"):
                return _BLUE
            return MUTED

        # aircraft family per airframe row; spare engines; spare engine types
        _fammap = fleet.set_index("Tail")["Aircraft_Family"].to_dict() if not fleet.empty else {}
        u["Aircraft_Family"] = [(_fammap.get(uid) if k == "AIRFRAME" else None)
                                for uid, k in zip(u["Unit_ID"], u["Unit_Kind"])]
        _installed = set()
        for _c in ("LHENG_ESN", "RHENG_ESN"):
            if _c in fleet.columns:
                _installed |= set(fleet[_c].dropna().astype(str))
        u["_spare"] = (u["Unit_Kind"] == "ENGINE") & (~u["Unit_ID"].astype(str).isin(_installed))
        _fams = sorted([f for f in u["Aircraft_Family"].dropna().unique()])

        # ── Monthly utilisation — real FH + FH/FC, with flexible filters ──
        st.markdown("#### Monthly utilisation")
        cl, cp = st.columns([1, 1.4])
        level = cl.selectbox("View by", ["Aircraft type", "Individual unit", "Spare engines"],
                             key="util_level")
        if level == "Aircraft type":
            pick = cp.selectbox("Aircraft type", ["All aircraft"] + _fams, key="util_pick_type")
            # "All aircraft" = airframes that belong to the fleet (exclude test/non-fleet rows)
            sub = u[(u["Unit_Kind"] == "AIRFRAME") & u["Aircraft_Family"].notna()] \
                if pick == "All aircraft" else u[u["Aircraft_Family"] == pick]
            label = pick
        elif level == "Spare engines":
            spares = sorted(u.loc[u["_spare"], "Unit_ID"].astype(str).unique())
            pick = cp.selectbox("Spare engine", ["All spare engines"] + spares, key="util_pick_spare")
            sub = u[u["_spare"]] if pick == "All spare engines" \
                else u[u["Unit_ID"].astype(str) == pick]
            label = pick
        else:
            units = sorted(u["Unit_ID"].astype(str).unique())
            pick = cp.selectbox("Unit  (tail / engine ESN / APU s/n)", units, key="util_pick_unit")
            sub = u[u["Unit_ID"].astype(str) == pick]
            label = pick

        if sub.empty:
            st.info("No utilisation rows for this selection.")
        else:
            ann_fh = float(sub["Annual_FH"].mean())
            ann_fc = float(sub["Annual_FC"].mean())
            ratio = (ann_fh / ann_fc) if ann_fc else float("nan")
            _agg = "average per unit" if len(sub) > 1 else "this unit"
            st.markdown(
                '<div class="kpi-row">'
                + _tile(f"{ann_fh:,.0f}", "Annual FH", _agg, ACCENT)
                + _tile(f"{ann_fc:,.0f}", "Annual FC", _agg, ACCENT2)
                + _tile(f"{ratio:.2f}" if ratio == ratio else "—", "FH : FC ratio",
                        "flight hours per cycle", C_HARD)
                + _tile(f"{len(sub):,}", "Units in selection", "", C_SOFT)
                + '</div>', unsafe_allow_html=True)

            _w = sub[_MONTHS].apply(pd.to_numeric, errors="coerce")
            fh_m = _w.multiply(sub["Annual_FH"].values, axis=0).mean()   # mean monthly FH
            fc_m = _w.multiply(sub["Annual_FC"].values, axis=0).mean()   # mean monthly FC
            ratio_m = (fh_m / fc_m.replace(0, pd.NA)).fillna(0)
            labels = [f"{fh:,.0f} FH<br><span style='color:#8fb4ff'>{r:.2f} FH/FC</span>"
                      for fh, r in zip(fh_m.values, ratio_m.values)]
            figm = px.bar(x=_MONTHS, y=fh_m.values, color_discrete_sequence=[ACCENT])
            figm.update_traces(text=labels, textposition="outside", cliponaxis=False,
                               textfont=dict(color=TEXT, size=12), marker_line_width=0,
                               hovertemplate="%{x}: %{y:,.0f} FH<extra></extra>")
            _mx = float(fh_m.max()) if len(fh_m) and pd.notna(fh_m.max()) else 10
            figm.update_layout(xaxis_title="Month", yaxis_title="Flight hours (FH) / month",
                               bargap=.3)
            figm.update_yaxes(range=[0, _mx * 1.3])
            st.plotly_chart(_sty(figm, f"Monthly flight hours — {label}  ({len(sub)} units)", h=400),
                            use_container_width=True)
            st.caption("Left axis = **actual monthly flight hours** (annual FH spread by the "
                       "seasonality weights). On each bar: FH that month + the FH/FC ratio.")

        # ── Fleet composition (active at a chosen date) ──
        st.markdown("#### Fleet composition")
        fl = fleet.copy()
        for c in ("Delivery_Date", "Redelivery_Date", "Extension_Date"):
            if c in fl.columns:
                fl[c] = pd.to_datetime(fl[c], errors="coerce")
        cda, _ = st.columns([1, 3])
        asof = cda.date_input("Active fleet as of", value=pd.Timestamp.today().date(),
                              key="fleet_asof",
                              help="Shows aircraft in service on this date: Delivery ≤ date ≤ "
                                   "Redelivery (or Extension, if the lease was extended).")
        _d = pd.Timestamp(asof)
        _start = fl["Delivery_Date"] if "Delivery_Date" in fl.columns else None
        _end = None
        if "Redelivery_Date" in fl.columns:
            _end = fl["Redelivery_Date"]
        if "Extension_Date" in fl.columns:
            _end = fl["Extension_Date"].fillna(_end) if _end is not None else fl["Extension_Date"]
        mask = pd.Series(True, index=fl.index)
        if _start is not None:
            mask &= _start.isna() | (_start <= _d)
        if _end is not None:
            mask &= _end.isna() | (_end >= _d)
        active = fl[mask]
        cmap = {f: _fam_color(f) for f in fleet["Aircraft_Family"].dropna().unique()}
        st.caption(f"**{len(active)}** aircraft active on {_d:%d %b %Y}  "
                   f"(of {len(fl)} in the database).")

        _, dcm, _ = st.columns([1, 2, 1])
        with dcm:
            if "Leased_Owned" in active.columns and not active.empty:
                lo = active["Leased_Owned"].value_counts().reset_index()
                lo.columns = ["Type", "Count"]
                figl = px.pie(lo, names="Type", values="Count", hole=.64, color="Type",
                              color_discrete_map={"Leased": ACCENT, "Owned": _GREEN})
                figl.update_traces(textinfo="label+percent", textfont_color="#fff", textfont_size=16,
                                   insidetextorientation="horizontal", pull=[0.02, 0.02],
                                   sort=False, rotation=90, domain=dict(x=[0.12, 0.88], y=[0, 1]),
                                   marker=dict(line=dict(color=BG0, width=3)),
                                   hovertemplate="%{label}: %{value} aircraft  (%{percent})<extra></extra>")
                figl = _sty(figl, "Owned vs leased", h=420, legend=True)
                figl.update_layout(legend=dict(orientation="h", y=-0.02, x=0.5, xanchor="center",
                                               title=""))
                figl.add_annotation(text=f"<b>{len(active)}</b><br>"
                                         f"<span style='font-size:12px;color:#8A93A6'>aircraft</span>",
                                    showarrow=False, x=0.5, y=0.5, xref="paper", yref="paper",
                                    font=dict(color="#fff", size=26))
                st.plotly_chart(figl, use_container_width=True)

        # aircraft by type — full width, taller, number labels never clipped
        if "Aircraft_Family" in active.columns and not active.empty:
            fam = active["Aircraft_Family"].value_counts().reset_index()
            fam.columns = ["Family", "Count"]
            figf = px.bar(fam, x="Family", y="Count", text="Count")
            figf.update_traces(marker_color=[cmap.get(f, ACCENT) for f in fam["Family"]],
                               textposition="outside", textfont=dict(color=TEXT, size=16),
                               marker_line_width=0, cliponaxis=False, width=0.62,
                               hovertemplate="<b>%{x}</b>: %{y} aircraft<extra></extra>")
            figf.update_layout(xaxis_title="", yaxis_title="aircraft", bargap=.35)
            figf.update_yaxes(range=[0, float(fam["Count"].max()) * 1.18])
            figf.update_xaxes(tickangle=-20)
            st.plotly_chart(_sty(figf, "Aircraft by type", h=440), use_container_width=True)

        # lessors / owners — stacked by aircraft type (one colour per type)
        if "Owner" in active.columns and not active.empty:
            ot = active.groupby(["Owner", "Aircraft_Family"]).size().reset_index(name="Count")
            totals = active.groupby("Owner").size()
            order = totals.sort_values(ascending=False).index.tolist()
            figo = px.bar(ot, x="Count", y="Owner", color="Aircraft_Family", orientation="h",
                          color_discrete_map=cmap, category_orders={"Owner": order[::-1]},
                          text="Count")
            figo.update_traces(marker_line_width=0, textposition="inside",
                               insidetextanchor="middle", texttemplate="%{text}",
                               textfont=dict(color="#fff", size=11),
                               hovertemplate="<b>%{y}</b> · %{fullData.name}: %{x}<extra></extra>")
            # total count at the end of each owner's bar
            for _own, _t in totals.items():
                figo.add_annotation(x=_t, y=_own, text=f"<b>{_t}</b>", showarrow=False,
                                    xanchor="left", xshift=8, font=dict(color=TEXT, size=13))
            figo.update_layout(barmode="stack", xaxis_title="aircraft", yaxis_title="",
                               legend_title="Type", bargap=.3)
            figo.update_xaxes(range=[0, float(totals.max()) * 1.12])
            st.plotly_chart(_sty(figo, "Aircraft per owner / lessor — by type",
                                 h=max(340, len(order) * 32 + 130), legend=True),
                            use_container_width=True)

        # ── Projected AOG (units with zero projected utilisation) ──
        st.markdown("#### Projected AOG — when units are grounded (no flying)")
        st.caption("Each bar is the **window a unit is projected to be grounded** (zero flight "
                   "hours) — e.g. an AOG lead-in or an out-of-service gap. Longer bar = more "
                   "months on the ground. Hover for the exact period.")
        frames = proj.aircraft_frames or {}
        if not frames:
            st.info("Run the projection to see projected AOG.")
        else:
            aog_h = st.slider("Look-ahead (months)", 6, min(60, len(proj.dates)), 24, key="aog_h")
            _a0 = proj.dates[0]
            _cut = _a0 + pd.offsets.MonthEnd(aog_h)
            _esn = {}
            if not fleet.empty:
                for _, _r in fleet.iterrows():
                    _esn[(_r["Tail"], "LHENG")] = _r.get("LHENG_ESN")
                    _esn[(_r["Tail"], "RHENG")] = _r.get("RHENG_ESN")

            def _span(mask, dates):
                z = dates[mask]
                start = z.min().to_period("M").to_timestamp()
                return start, z.max(), int(mask.sum())

            ac_rows, eng_rows = [], []
            for _tail, _fr in frames.items():
                _fr = _fr.copy()
                _fr["DATE"] = pd.to_datetime(_fr["DATE"])
                _nx = _fr[(_fr["DATE"] >= _a0) & (_fr["DATE"] <= _cut)]
                if _nx.empty:
                    continue
                if "REG-UTIL FH" in _nx.columns:
                    _zm = pd.to_numeric(_nx["REG-UTIL FH"], errors="coerce").fillna(0) == 0
                    if _zm.any():
                        s, e, n = _span(_zm, _nx["DATE"])
                        ac_rows.append({"Tail": _tail, "Start": s, "End": e, "months": n,
                                        "From": s.strftime("%b %Y"), "To": e.strftime("%b %Y")})
                for _pos in ("LHENG", "RHENG"):
                    _col = f"{_pos}-UTIL FH"
                    if _col in _nx.columns:
                        _em = pd.to_numeric(_nx[_col], errors="coerce").fillna(0) == 0
                        if _em.any():
                            s, e, n = _span(_em, _nx["DATE"])
                            _sn = _esn.get((_tail, _pos)) or f"{_tail}-{_pos}"
                            eng_rows.append({"Engine": str(_sn), "Start": s, "End": e, "months": n,
                                             "From": s.strftime("%b %Y"), "To": e.strftime("%b %Y")})

            AMBER, REDS = "#F6A609", "#E06666"
            st.markdown(
                '<div class="kpi-row">'
                + _tile(f"{len(ac_rows)}", "Aircraft grounded", f"in next {aog_h} months", AMBER)
                + _tile(f"{len({r['Engine'] for r in eng_rows})}", "Engines grounded",
                        f"in next {aog_h} months", REDS)
                + _tile(f"{sum(r['months'] for r in ac_rows):,}", "Aircraft-months", "on the ground", C_SOFT)
                + _tile(f"{len(frames) - len(ac_rows)}", "Flying throughout", "no AOG", C_HARD)
                + '</div>', unsafe_allow_html=True)

            def _aog_timeline(rows, ycol, title, color):
                df = pd.DataFrame(rows).sort_values("Start")
                fig = px.timeline(df, x_start="Start", x_end="End", y=ycol,
                                  color_discrete_sequence=[color], custom_data=["months", "From", "To"])
                fig.update_yaxes(autorange="reversed", title="")
                fig.update_traces(marker_line_width=0, opacity=.95,
                                  hovertemplate="<b>%{y}</b><br>%{customdata[0]} months grounded"
                                                "<br>%{customdata[1]} → %{customdata[2]}<extra></extra>")
                fig.update_layout(xaxis_title="")
                return _sty(fig, title, h=max(240, len(df) * 46 + 110))

            ca, ce = st.columns(2)
            with ca:
                if ac_rows:
                    st.plotly_chart(_aog_timeline(ac_rows, "Tail", "Aircraft — grounded period", AMBER),
                                    use_container_width=True)
                else:
                    st.success(f"✈️ No aircraft projected AOG in the next {aog_h} months.")
            with ce:
                if eng_rows:
                    st.plotly_chart(_aog_timeline(eng_rows, "Engine", "Engines — grounded period", REDS),
                                    use_container_width=True)
                else:
                    st.success(f"🔧 No engines projected AOG in the next {aog_h} months.")

        st.download_button("⬇ Download utilisation assumptions (CSV)",
                           util.to_csv(index=False).encode(), "utilisation.csv", "text/csv")

# ── Current Fleet Status (versatile live monitor) ─────────────
if nav == NAV_STATUS:
    st.subheader("Current fleet status — live monitor")
    _us = ds.get("Unit_Status")
    _comp = ds.get("Components")
    _evdef = ds.get("Events")
    if _us is None or _us.empty:
        st.info("No status data in the database.")
    else:
        _info = {}
        if _comp is not None and not _comp.empty:
            for _, r in _comp.iterrows():
                _info[str(r["Component_ID"])] = (str(r["Unit_Kind"]), str(r["Unit_Type"]),
                                                 str(r.get("Position") or ""),
                                                 str(r.get("Installed_On_Tail") or ""))
        if not fleet.empty:
            for _, r in fleet.iterrows():
                _info[str(r["Tail"])] = ("AIRFRAME", str(r.get("Aircraft_Family") or ""),
                                         "AIRFRAME", str(r["Tail"]))
        _lim = {}
        if _evdef is not None and not _evdef.empty:
            for _, r in _evdef.iterrows():
                _lim[(str(r["Unit_Type"]), str(r["Event_ID"]))] = (r.get("Hard_Limit"),
                                                                   str(r.get("Driver") or ""))
        _POSL = {"LHENG": "LH", "RHENG": "RH", "APU": "APU", "LHPROP": "LH", "RHPROP": "RH",
                 "AIRFRAME": "—", "SPARE": "Spare"}
        _rows = []
        for _, r in _us.iterrows():
            uid = str(r["Unit_ID"])
            inf = _info.get(uid)
            if not inf:
                continue
            kind, utype, pos, tail = inf
            evid = str(r["Event_ID"])
            hl, drv = _lim.get((utype, evid), (None, str(r.get("Driver") or "")))
            drv = drv or str(r.get("Driver") or "")
            # current + at-last-SV readings
            tsn = pd.to_numeric(r.get("TSN"), errors="coerce")
            csn = pd.to_numeric(r.get("CSN"), errors="coerce")
            tsn_l = pd.to_numeric(r.get("TSN_At_LSV"), errors="coerce")
            csn_l = pd.to_numeric(r.get("CSN_At_LSV"), errors="coerce")
            slsv_fh = pd.to_numeric(r.get("SLSV_FH"), errors="coerce")
            slsv_fc = pd.to_numeric(r.get("SLSV_FC"), errors="coerce")
            # cycles since last SV = CSN now − CSN @ last SV; but the @-LSV reading is
            # engine-level — when it's missing/zero, trust the per-event AMOS value.
            cyc_sv = (csn - csn_l) if (pd.notna(csn) and pd.notna(csn_l) and csn_l) else slsv_fc
            # "since last event" in the event's own driver units
            if drv.startswith("CAL") or drv == "CLND":
                consumed = pd.to_numeric(r.get("SLSV_Months"), errors="coerce")
            elif drv == "FC":
                consumed = cyc_sv
            else:   # FH — hours since last event = TSN now − TSN @ last SV (when valid)
                consumed = (tsn - tsn_l) if (pd.notna(tsn) and pd.notna(tsn_l) and tsn_l) else slsv_fh
            hln = pd.to_numeric(hl, errors="coerce")
            remaining = (hln - consumed) if (pd.notna(hln) and pd.notna(consumed)) else None
            pct = round(consumed / hln * 100, 1) if (pd.notna(hln) and hln and pd.notna(consumed)) else None

            def _r0(x):
                return round(x, 0) if pd.notna(x) else None

            def _r1(x):
                return round(x, 1) if pd.notna(x) else None

            _rows.append({"Unit": uid, "Kind": kind, "Type": utype, "Pos": _POSL.get(pos, pos),
                          "Tail": tail, "Event": evid, "Driver": drv,
                          "Consumed": _r1(consumed), "CycSV": _r0(cyc_sv),
                          "Limit": hln if pd.notna(hln) else None,
                          "Remaining": _r1(remaining), "% used": pct,
                          "TSN_LSV": _r0(tsn_l), "CSN_LSV": _r0(csn_l),
                          "TSN_now": _r0(tsn), "CSN_now": _r0(csn),
                          "Last_SV": r.get("Last_SV_Date")})
        sdf = pd.DataFrame(_rows)
        if sdf.empty:
            st.info("No matching status rows.")
        else:
            def _prune(key, opts):
                """Drop stale selections so options can cascade without Streamlit errors."""
                if key in st.session_state:
                    st.session_state[key] = [x for x in st.session_state[key] if x in opts]

            c1, c2, c3, c4 = st.columns(4)
            # Cascading filters — each dropdown only offers values relevant to the
            # selections above it (e.g. pick AIRFRAME → Tail list shows aircraft tails).
            o_kind = sorted(sdf["Kind"].dropna().unique())
            _prune("st_kind", o_kind)
            _k = c1.multiselect("Unit kind", o_kind, key="st_kind", placeholder="All kinds")
            f1 = sdf[sdf["Kind"].isin(_k)] if _k else sdf

            o_type = sorted(f1["Type"].dropna().unique())
            _prune("st_type", o_type)
            _t = c2.multiselect("Type", o_type, key="st_type", placeholder="All types")
            f2 = f1[f1["Type"].isin(_t)] if _t else f1

            o_ev = sorted(f2["Event"].dropna().unique())
            _prune("st_ev", o_ev)
            _e = c3.multiselect("Event", o_ev, key="st_ev", placeholder="All events")
            f3 = f2[f2["Event"].isin(_e)] if _e else f2

            o_tu = sorted(set(f3["Tail"].dropna()) | set(f3["Unit"].dropna()))
            _prune("st_tail", o_tu)
            _tl = c4.multiselect("Tail / unit", o_tu, key="st_tail", placeholder="All")
            v = f3[f3["Tail"].isin(_tl) | f3["Unit"].isin(_tl)] if _tl else f3
            st.caption(f"{len(v):,} unit × event rows (as of the snapshot).")
            v = v.sort_values("% used", ascending=False, na_position="last")
            st.markdown(_pretty_table(
                v, ["Tail", "Unit", "Kind", "Type", "Pos", "Event", "Driver",
                    "Consumed", "CycSV", "Limit", "Remaining", "% used",
                    "TSN_LSV", "CSN_LSV", "TSN_now", "CSN_now", "Last_SV"],
                height=520, cap=800, signed_cols=["Remaining"],
                labels={"Unit": "Unit / serial", "Consumed": "Since last event",
                        "CycSV": "Cycles since SV", "% used": "% of limit",
                        "TSN_LSV": "TSN @ last SV", "CSN_LSV": "CSN @ last SV",
                        "TSN_now": "TSN now", "CSN_now": "CSN now", "Last_SV": "Last SV"}),
                unsafe_allow_html=True)
            st.write("")
            st.download_button("⬇ Download current status (CSV)", v.to_csv(index=False).encode(),
                               "current_fleet_status.csv", "text/csv")

# ── Powerplant Components (engines & APUs, incl. spares) ──────
if nav == NAV_POWER:
    st.subheader("Powerplant components — engines & APUs")
    comp = ds.get("Components")
    us = ds.get("Unit_Status")
    llp = ds.get("LLP_Registry")
    if comp.empty:
        st.info("No components in the database.")
    else:
        pw = comp[comp["Unit_Kind"].isin(["ENGINE", "APU"])].copy()
        _POS = {"LHENG": "LH", "RHENG": "RH", "APU": "APU", "SPARE": "Spare"}
        pw["Pos"] = pw["Position"].map(_POS).fillna(pw["Position"].astype(str))
        pw["State"] = pw["Position"].apply(lambda p: "Spare" if str(p) == "SPARE" else "Installed")

        # current status per unit (last event) — straight from the DB, no hardcoded event names
        if us is not None and not us.empty:
            _aggcols = {"TSN": ("TSN", "first"), "CSN": ("CSN", "first"),
                        "Last_SV": ("Last_SV_Date", "first")}
            if "TSN_At_LSV" in us.columns:
                _aggcols["TSN_at_SV"] = ("TSN_At_LSV", "first")
            if "WS_Type" in us.columns:
                _aggcols["WS"] = ("WS_Type", "first")
            _agg = (us.groupby(us["Unit_ID"].astype(str)).agg(**_aggcols)
                    .reset_index().rename(columns={"Unit_ID": "_cid"}))
            pw = pw.merge(_agg, left_on=pw["Component_ID"].astype(str), right_on="_cid", how="left")
        for _c in ("TSN", "CSN", "Last_SV", "TSN_at_SV", "WS"):
            if _c not in pw.columns:
                pw[_c] = None
        # LLP count per unit
        if llp is not None and not llp.empty:
            _lc = llp.groupby(llp["Unit_ID"].astype(str)).size().rename("LLPs")
            pw["LLPs"] = pw["Component_ID"].astype(str).map(_lc).fillna(0).astype(int)
        else:
            pw["LLPs"] = 0

        _eng = pw[pw["Unit_Kind"] == "ENGINE"]
        _apu = pw[pw["Unit_Kind"] == "APU"]
        st.markdown(
            '<div class="kpi-row">'
            + _tile(f"{(_eng['State'] == 'Installed').sum()}", "Installed engines", "on-wing", ACCENT)
            + _tile(f"{(_eng['State'] == 'Spare').sum()}", "Spare engines", "separate units", C_SOFT)
            + _tile(f"{(_apu['State'] == 'Installed').sum()}", "Installed APUs", "on-wing", ACCENT2)
            + _tile(f"{(_apu['State'] == 'Spare').sum()}", "Spare APUs", "separate units", C_HARD)
            + '</div>', unsafe_allow_html=True)

        f1, f2, f3 = st.columns([1, 1, 1])
        _kinds = f1.multiselect("Kind", ["ENGINE", "APU"], default=["ENGINE", "APU"], key="pw_kind")
        _types = sorted(pw["Unit_Type"].dropna().astype(str).unique())
        _typ = f2.multiselect("Type", _types, key="pw_type", placeholder="All types")
        _state = f3.multiselect("State", ["Installed", "Spare"], key="pw_state",
                                placeholder="Installed + Spare")
        view = pw.copy()
        if _kinds:
            view = view[view["Unit_Kind"].isin(_kinds)]
        if _typ:
            view = view[view["Unit_Type"].astype(str).isin(_typ)]
        if _state:
            view = view[view["State"].isin(_state)]

        view = view.rename(columns={"Component_ID": "Serial", "Unit_Type": "Type",
                                    "Installed_On_Tail": "On_Tail", "TSN_at_SV": "TSN_at_LSV",
                                    "Last_SV": "Last_SV_date", "WS": "Last_WS"})
        st.caption(f"{len(view)} powerplant units.  **TSN/CSN** = current (as of the snapshot date); "
                   "**TSN@LSV** = TSN at the last shop visit; **Last SV / WS** from Fleet Maintenance Status.")
        st.markdown(_pretty_table(view, ["Serial", "Type", "Pos", "State", "On_Tail",
                                         "TSN", "TSN_at_LSV", "CSN", "Last_SV_date", "Last_WS",
                                         "LLPs"],
                                  height=420, cap=400), unsafe_allow_html=True)
        st.write("")

        # ── drill-down: one unit's sub-events + LLPs ──
        st.markdown("#### Unit detail — shop visits & LLPs")
        _units = sorted(view["Serial"].astype(str).unique())
        if _units:
            sel = st.selectbox("Select a unit (engine ESN / APU s/n)", _units, key="pw_unit")
            # event display names come straight from the database (Events sheet) — not the code
            _evdf = ds.get("Events")
            _EVN = (dict(zip(_evdf["Event_ID"].astype(str), _evdf["Event_Name"].astype(str)))
                    if _evdf is not None and not _evdf.empty else {})
            # events/counters for this unit
            su = us[us["Unit_ID"].astype(str) == sel].copy() if (us is not None and not us.empty) else pd.DataFrame()
            cda, cdb = st.columns(2)
            with cda:
                st.markdown("**Shop-visit counters (since last SV)**")
                if not su.empty:
                    su["Event"] = su["Event_ID"].astype(str).map(_EVN).fillna(su["Event_ID"].astype(str))
                    su = su.rename(columns={"SLSV_FH": "FH since SV", "SLSV_FC": "FC since SV",
                                            "Last_SV_Date": "Last SV"})
                    st.markdown(_pretty_table(su, ["Event", "Driver", "FH since SV", "FC since SV",
                                                   "TSN", "CSN", "Last SV"], height=240),
                                unsafe_allow_html=True)
                else:
                    st.caption("No counter rows for this unit.")
            with cdb:
                st.markdown("**Life-limited parts (LLPs)**")
                sl = llp[llp["Unit_ID"].astype(str) == sel].copy() if (llp is not None and not llp.empty) else pd.DataFrame()
                if not sl.empty:
                    sl = sl.sort_values("ToGo")
                    st.markdown(_pretty_table(sl, ["LLP_PN", "Description", "Limit_Cycles",
                                                   "SLSV", "ToGo"], height=240, cap=100),
                                unsafe_allow_html=True)
                    st.caption(f"{len(sl)} LLPs on this unit (nearest to limit first).")
                else:
                    st.caption("No LLPs recorded for this unit.")

# ── Edit Data (edit any sheet → download updated database) ────
if nav == NAV_EDIT:
    import io as _io
    st.subheader("Edit the database & download the updated file")
    st.caption("Change any value directly here, then **download the updated database**. "
               "Click **Apply & re-run** to project with your edits straight away — or re-upload "
               "the downloaded file later.")

    # load every sheet fresh from the uploaded bytes (once per uploaded file)
    if st.session_state.get("_edit_src") != _dbhash:
        try:
            _xls = pd.ExcelFile(_io.BytesIO(db_bytes))
            st.session_state["edit_sheets"] = {s: pd.read_excel(_xls, s)
                                               for s in _xls.sheet_names if not str(s).startswith("_")}
            st.session_state["_edit_src"] = _dbhash
        except Exception as exc:
            st.error(f"Could not open the workbook for editing: {exc}")
            st.stop()

    _sheets = st.session_state.get("edit_sheets", {})
    if not _sheets:
        st.info("No editable sheets found.")
    else:
        _names = list(_sheets.keys())
        c1, c2 = st.columns([1.3, 2])
        _pick = c1.selectbox("Sheet to edit", _names,
                             index=_names.index("Fleet") if "Fleet" in _names else 0,
                             key="edit_pick")
        c2.caption(f"**{_pick}** — {len(_sheets[_pick]):,} rows × {_sheets[_pick].shape[1]} columns. "
                   "Edit cells, add rows (＋ at the bottom), or delete rows (select + ⌫).")
        _edited = st.data_editor(_sheets[_pick], num_rows="dynamic", use_container_width=True,
                                 height=460, key=f"ed_{_pick}")
        _sheets[_pick] = _edited          # keep this sheet's edits in session

        # build the updated workbook from all (possibly edited) sheets
        def _build_xlsx(sheets_dict):
            buf = _io.BytesIO()
            with pd.ExcelWriter(buf, engine="xlsxwriter") as _w:
                for _s, _df in sheets_dict.items():
                    _df.to_excel(_w, sheet_name=str(_s)[:31], index=False)
            return buf.getvalue()

        st.divider()
        b1, b2, b3 = st.columns([1.3, 1.1, 2])
        _xlsx_bytes = _build_xlsx(_sheets)
        b1.download_button("⬇ Download updated database (.xlsx)", _xlsx_bytes,
                           "FLEET_MAINTENANCE_STATUS_DATABASE.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True)
        if b2.button("✅ Apply & re-run", type="primary", use_container_width=True):
            st.session_state["edited_db_bytes"] = _xlsx_bytes
            _run.clear()
            st.session_state.pop("_dbhash", None)
            st.rerun()
        b3.caption("Download keeps a copy for later · Apply & re-run projects with your edits now.")

# ── Export ────────────────────────────────────────────────────
if nav == NAV_EXP:
    st.subheader("Download the full projection")
    st.caption("A structured workbook: SV_PLAN_v1 (calendar), SV_PLAN_v2 (event × aircraft), "
               "and one sheet per aircraft with the month-by-month counters.")
    if st.button("🛠 Build projection workbook", type="primary"):
        with st.spinner("Building workbook (one sheet per aircraft)…"):
            st.session_state["workbook"] = build_workbook(proj)
    if st.session_state.get("workbook"):
        st.download_button("⬇ Download FLEET_MAINTENANCE_PROJECTION.xlsx",
                           st.session_state["workbook"],
                           "FLEET_MAINTENANCE_PROJECTION.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.divider()
    if proj.sv_plan_v1 is not None:
        st.download_button("⬇ SV calendar (CSV)", proj.sv_plan_v1.to_csv(index=False).encode(),
                           "sv_plan_v1.csv", "text/csv")
    if proj.sv_plan_v2 is not None:
        st.download_button("⬇ SV matrix (CSV)", proj.sv_plan_v2.to_csv(index=False).encode(),
                           "sv_plan_v2.csv", "text/csv")

# ── Guide & Templates (onboarding: sample file + 2-page overview) ──
def _overview_html() -> str:
    """Self-contained, print-to-PDF-ready 2-page overview of the app + data model."""
    A = "#1a73c7"
    return f"""<!doctype html><html><head><meta charset="utf-8">
<title>Fleet Forecast — Overview</title><style>
@page {{ size: A4; margin: 16mm; }}
* {{ box-sizing: border-box; }}
body {{ font-family: -apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif; color:#1b2330;
       margin:0; font-size:12px; line-height:1.5; }}
.page {{ page-break-after: always; padding:0 4px; }}
.page:last-child {{ page-break-after: auto; }}
h1 {{ font-size:26px; margin:0 0 2px; letter-spacing:-.5px; }}
h1 .b {{ color:{A}; }}
.kick {{ color:#6b7687; font-size:12px; text-transform:uppercase; letter-spacing:2px; margin-bottom:14px; }}
h2 {{ font-size:14px; color:{A}; margin:18px 0 6px; border-bottom:2px solid {A}22; padding-bottom:3px; }}
.lead {{ font-size:13.5px; color:#39424f; margin:10px 0 4px; }}
.grid {{ display:grid; grid-template-columns:1fr 1fr; gap:10px 22px; margin-top:6px; }}
.card {{ border:1px solid #e2e7ee; border-left:3px solid {A}; border-radius:6px; padding:8px 11px; }}
.card b {{ color:#141b26; }}
table {{ width:100%; border-collapse:collapse; margin-top:4px; font-size:11px; }}
th {{ text-align:left; background:{A}; color:#fff; padding:5px 8px; font-weight:600; }}
td {{ padding:5px 8px; border-bottom:1px solid #e8ecf2; vertical-align:top; }}
tr:nth-child(even) td {{ background:#f6f8fb; }}
.mono {{ font-family:ui-monospace,Menlo,Consolas,monospace; background:#eef2f7; padding:1px 5px; border-radius:4px; }}
.foot {{ margin-top:16px; color:#8a93a3; font-size:10px; border-top:1px solid #e2e7ee; padding-top:6px; }}
.pill {{ display:inline-block; background:{A}15; color:{A}; border-radius:999px; padding:1px 9px; font-size:10px; font-weight:600; }}
</style></head><body>

<div class="page">
  <div class="kick">AEGEAN · Technical Fleet Planning</div>
  <h1>Fleet <span class="b">Forecast</span> — Overview</h1>
  <p class="lead">A single, file-driven engine that projects every airframe, engine, APU and
  propeller shop visit across the lease horizon. One Excel workbook is the master source:
  change a limit, a reading or a date there and the whole forecast updates — nothing is
  hard-coded in the application.</p>

  <h2>What it does</h2>
  <div class="grid">
    <div class="card"><b>Projects shop visits</b><br>Every maintenance event (checks, engine SVs,
      APU, propeller, LLPs) is projected month-by-month to its due date using each event's own
      driver and limit.</div>
    <div class="card"><b>Drivers &amp; links</b><br>Each event runs on Flight Hours, Flight Cycles
      or calendar months, with soft/hard links so events bundled into one visit are recognised.</div>
    <div class="card"><b>Live fleet status</b><br>A snapshot monitor shows, per unit, time since the
      last event, limit, remaining life (green) or overdue (red), and current TSN/CSN.</div>
    <div class="card"><b>Powerplant &amp; LLPs</b><br>Engines and APUs — installed and spare —
      with TSN/CSN at last SV, workscope class and life-limited-part registries.</div>
  </div>

  <h2>How it works</h2>
  <table>
    <tr><th>Step</th><th>Detail</th></tr>
    <tr><td><b>1 · Upload</b></td><td>Upload the master database workbook. The app starts empty —
      your data never ships with it.</td></tr>
    <tr><td><b>2 · Snapshot</b></td><td>The <span class="mono">As_Of_Date</span> in Config is the
      reference point; all current readings (TSN / CSN / time-since-last-SV) are as of that date.</td></tr>
    <tr><td><b>3 · Project</b></td><td>The engine walks each aircraft and its installed units to the
      end of its lease (redelivery or extension), emitting every due event.</td></tr>
    <tr><td><b>4 · Explore &amp; export</b></td><td>Dashboard, calendar, powerplant and status views;
      export the full month-by-month projection workbook.</td></tr>
  </table>

  <h2>Sections at a glance</h2>
  <p><span class="pill">Current Fleet Status</span> <span class="pill">SV Dashboard</span>
  <span class="pill">Fleet &amp; Utilisation</span> <span class="pill">SV Calendar</span>
  <span class="pill">Event Log</span> <span class="pill">LLP Status</span>
  <span class="pill">Aircraft Detail</span> <span class="pill">Powerplant</span>
  <span class="pill">Edit Data</span> <span class="pill">Export</span></p>
  <div class="foot">Fleet Forecast · file-driven maintenance projection · page 1 of 2</div>
</div>

<div class="page">
  <div class="kick">Database Guide</div>
  <h1>Filling the <span class="b">master file</span></h1>
  <p class="lead">The workbook is relational — identifiers (tails and serials) must match across
  sheets. Keep the generic definition sheets (Events, Unit_Types, Event_Links) and fill the
  fleet-specific sheets with your data.</p>

  <table>
    <tr><th>Sheet</th><th>What to put in it</th></tr>
    <tr><td><b>Config</b></td><td>Snapshot / reference date (<span class="mono">As_Of_Date</span>).</td></tr>
    <tr><td><b>Unit_Types</b></td><td>Every type: airframe families, engine types, APU, propeller,
      plus each type's LLP driver &amp; linked event.</td></tr>
    <tr><td><b>Events</b></td><td>Each maintenance event with its <b>Driver</b> (FH / FC /
      CALENDAR_MONTHS), <b>Hard_Limit</b>, soft threshold and workscope class.</td></tr>
    <tr><td><b>Event_Links</b></td><td>Soft / hard links — events performed together in one visit.</td></tr>
    <tr><td><b>Fleet</b></td><td>Aircraft: tail, family, delivery / redelivery / extension dates,
      installed engine &amp; APU serials.</td></tr>
    <tr><td><b>Components</b></td><td>Engines, APUs, propellers — serial, type, position,
      installed-on tail (blank = spare).</td></tr>
    <tr><td><b>Unit_Status</b></td><td>Live readings per unit × event: TSN, CSN, TSN/CSN @ last SV,
      time since last SV, last-SV date.</td></tr>
    <tr><td><b>Utilization</b></td><td>Annual FH / FC and monthly distribution per unit.</td></tr>
    <tr><td><b>LLP_Registry</b></td><td>Life-limited parts per engine: PN, SN, limit, cycles used.</td></tr>
    <tr><td><b>Financial_* / MR_Rates / …</b></td><td>Financial module inputs (rates, SV cost matrix,
      end-of-lease terms).</td></tr>
  </table>

  <h2>Golden rules</h2>
  <div class="grid">
    <div class="card"><b>IDs must match</b><br>A tail or serial written in Fleet must appear
      identically in Components, Unit_Status and Utilization.</div>
    <div class="card"><b>Leave unknowns blank</b><br>Missing last-SV dates stay empty — the engine
      falls back to time-since-last-SV; never invent a date.</div>
    <div class="card"><b>The file decides</b><br>Limits, drivers, names and links live in the file,
      not the code. Re-upload and everything recalculates.</div>
    <div class="card"><b>Spares are units too</b><br>An engine or APU with a blank tail is a spare —
      tracked as its own inventory item.</div>
  </div>

  <h2>Get started</h2>
  <p class="lead">Download the sample workbook from the app, open it to see a worked 3-aircraft
  example, replace the demo rows with your fleet, and upload it back. That's it.</p>
  <div class="foot">Fleet Forecast · database guide · page 2 of 2</div>
</div>
</body></html>"""


if nav == NAV_GUIDE:
    st.subheader("Guide & templates")
    st.caption("Everything someone new needs to start: a worked sample workbook and a printable "
               "2-page overview of the app and the data model.")

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("#### 📘 2-page overview")
        st.write("A concise, presentation-grade explanation of what the app does and how to fill "
                 "the database. Open it and **Print → Save as PDF** for a shareable deck.")
        st.download_button("⬇ Download overview (HTML, print-to-PDF)",
                           _overview_html().encode("utf-8"),
                           "Fleet_Forecast_Overview.html", "text/html",
                           use_container_width=True)
    with c2:
        st.markdown("#### 📗 Sample database")
        st.write("A ready-to-upload workbook with **3 demo aircraft** (random data) and the full "
                 "sheet structure. Download it, replace the demo rows with your fleet, and upload.")
        _sample_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                    "assets", "SAMPLE_fleet_database.xlsx")
        if os.path.exists(_sample_path):
            with open(_sample_path, "rb") as _sf:
                st.download_button("⬇ Download sample database (.xlsx)", _sf.read(),
                                   "SAMPLE_fleet_database.xlsx",
                                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   use_container_width=True)
        else:
            st.info("Sample file not bundled in this build.")

    st.divider()
    st.markdown("#### How the pieces fit together")
    st.markdown(
        "- **The master file is the single source of truth** — limits, drivers, event names and "
        "links all live there, never in the code.\n"
        "- **The app is empty until you upload** — your fleet data never ships with it.\n"
        "- **Identifiers must match across sheets** — a tail/serial in *Fleet* must appear "
        "identically in *Components*, *Unit_Status* and *Utilization*.\n"
        "- **Re-upload to recalculate** — change anything in the file and the whole projection updates."
    )

# ── Data notes ────────────────────────────────────────────────
if nav == NAV_NOTES:
    st.subheader("Ingestion notes")
    if not ds.warnings:
        st.success("No data warnings.")
    for w in ds.warnings:
        st.write("⚠️ ", w)
    st.caption("Soft-bundling runs on the 'years remaining' threshold. Financial calculations "
               "are not yet active in this build.")
