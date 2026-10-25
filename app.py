"""
Fleet ERP — web app
====================
Upload the master database workbook, optionally upload a fresh LLP report
(raw export), run the technical projection, and browse the results.
"""

from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from streamlit_option_menu import option_menu

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
NAV_EXP   = "Export"
NAV_NOTES = "Data Notes"
NAV_ITEMS = [NAV_ASK, NAV_UTIL, NAV_DASH, NAV_CAL, NAV_LOG, NAV_LLP, NAV_AC, NAV_EXP, NAV_NOTES]
NAV_ICONS = ["chat-dots", "graph-up-arrow", "bar-chart-fill", "calendar3", "list-task",
             "wrench-adjustable", "airplane", "box-arrow-down", "info-circle"]

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


# ── Sidebar: uploads + navigation ─────────────────────────────
with st.sidebar:
    st.header("Database")
    db_file = st.file_uploader("Fleet Maintenance Status Database (.xlsx)", type=["xlsx"], key="db")
    _DEFAULT_DB = Path(__file__).parent / "data" / "FLEET_MAINTENANCE_STATUS_DATABASE.xlsx"
    if db_file is None and _DEFAULT_DB.exists():
        st.caption("Using the bundled database — upload a file above to override.")

    st.write("")
    nav = option_menu(menu_title="Navigation", options=NAV_ITEMS, icons=NAV_ICONS,
                      menu_icon="grid-1x2-fill", default_index=0, styles=NAV_STYLES,
                      key="nav")

    st.divider()
    run_clicked = st.button("↻ Re-run projection", type="primary", use_container_width=True)

# resolve the database source: uploaded file, else the bundled default
db_bytes = db_file.getvalue() if db_file is not None else (
    _DEFAULT_DB.read_bytes() if _DEFAULT_DB.exists() else None)

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

# ── KPI tiles (Power BI style) ────────────────────────────────
_n_ac = fleet["Tail"].nunique() if not fleet.empty else 0
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
            f"<td>{_fmt(r.get('Projected_SV'))}</td><td>{_fmt(r.get('Reason'))}</td></tr>"
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


def _pretty_table(df: pd.DataFrame, columns: list[str], reason_col: str = "Reason",
                  height: int = 360, cap: int = 200) -> str:
    """Render a dark, professional HTML table: sticky header, zebra rows, hover,
    tabular numbers, and colour-coded 'Reason' pills."""
    cols = [c for c in columns if c in df.columns]
    if not cols:
        return ""
    d = df[cols].head(cap)
    numeric = {c for c in cols if pd.api.types.is_numeric_dtype(d[c])}

    def _lab(c):
        return str(c).replace("_", " ")

    def _numfmt(v):
        try:
            f = float(v)
        except (TypeError, ValueError):
            return _fmt(v)
        if pd.isna(f):
            return ""
        return f"{int(f):,}" if f == int(f) else f"{f:,.1f}"

    def _badge(val):
        s = _fmt(val)
        if not s:
            return ""
        low = s.lower()
        c = (C_HARD if "hard link" in low else
             C_SOFT if "soft link" in low else
             C_LLP if ("llp" in low or "replace" in low) else
             C_LIMIT)
        return (f'<span style="display:inline-block;padding:2px 11px;border-radius:999px;'
                f'background:{c}22;color:{c};border:1px solid {c}66;font-size:.72rem;'
                f'font-weight:600;white-space:nowrap">{s}</span>')

    def _cell(c, v):
        if c == reason_col:
            return f"<td>{_badge(v)}</td>"
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


# ── Global fleet cards (fleet composition only) ───────────────
_famc = fleet["Aircraft_Family"].value_counts() if not fleet.empty else pd.Series(dtype=int)
_util_g = ds.get("Utilization")
_inst_e, _inst_a = set(), set()
if not fleet.empty:
    for _c in ("LHENG_ESN", "RHENG_ESN"):
        if _c in fleet.columns:
            _inst_e |= set(fleet[_c].dropna().astype(str))
    if "APU_SN" in fleet.columns:
        _inst_a = set(fleet["APU_SN"].dropna().astype(str))
if _util_g is not None and not _util_g.empty:
    _eng_u = _util_g[_util_g["Unit_Kind"] == "ENGINE"]["Unit_ID"].astype(str)
    _apu_u = _util_g[_util_g["Unit_Kind"] == "APU"]["Unit_ID"].astype(str)
    _n_spare_eng = int((~_eng_u.isin(_inst_e)).sum())
    _n_spare_apu = int((~_apu_u.isin(_inst_a)).sum())
else:
    _n_spare_eng = _n_spare_apu = 0

st.markdown(
    '<div style="display:grid;grid-template-columns:0.9fr 1.7fr 1.2fr;gap:.8rem;margin:.2rem 0 1rem;">'
    + _tile(f"{_n_ac}", "Total aircraft", "in fleet", ACCENT)
    + _tile_break("Aircraft by type", list(_famc.items()), ACCENT2)
    + _tile_break("Spares", [("Spare engines", _n_spare_eng), ("Spare APUs", _n_spare_apu)],
                  C_SOFT, total=_n_spare_eng + _n_spare_apu)
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
        kind_label = {"AIRFRAME": "Airframe", "ENGINE": "Engine", "APU": "APU"}
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
        scope_sel = f1.multiselect("Category", ["Airframe", "Landing gear", "Engine", "APU", "LLP"],
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
    from fleet_erp.ui import assistant as _assistant
    import anthropic

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
            sub = u[u["Unit_Kind"] == "AIRFRAME"] if pick == "All aircraft" \
                else u[u["Aircraft_Family"] == pick]
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

        dc1, _dc2 = st.columns([1, 1.1])
        with dc1:
            if "Leased_Owned" in active.columns and not active.empty:
                lo = active["Leased_Owned"].value_counts().reset_index()
                lo.columns = ["Type", "Count"]
                figl = px.pie(lo, names="Type", values="Count", hole=.6, color="Type",
                              color_discrete_map={"Leased": ACCENT, "Owned": _GREEN})
                figl.update_traces(textinfo="label+percent", textfont_color="#fff", textfont_size=16,
                                   insidetextorientation="horizontal",
                                   marker=dict(line=dict(color=BG0, width=2)))
                st.plotly_chart(_sty(figl, "Owned vs leased", h=360, legend=True),
                                use_container_width=True)

        # aircraft by type — full width, taller, number labels never clipped
        if "Aircraft_Family" in active.columns and not active.empty:
            fam = active["Aircraft_Family"].value_counts().reset_index()
            fam.columns = ["Family", "Count"]
            figf = px.bar(fam, x="Family", y="Count", color="Family",
                          color_discrete_map=cmap, text="Count")
            figf.update_traces(textposition="outside", textfont=dict(color=TEXT, size=16),
                               marker_line_width=0, cliponaxis=False,
                               hovertemplate="<b>%{x}</b>: %{y} aircraft<extra></extra>")
            figf.update_layout(xaxis_title="", yaxis_title="aircraft", bargap=.28)
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

# ── Data notes ────────────────────────────────────────────────
if nav == NAV_NOTES:
    st.subheader("Ingestion notes")
    if not ds.warnings:
        st.success("No data warnings.")
    for w in ds.warnings:
        st.write("⚠️ ", w)
    st.caption("Soft-bundling runs on the 'years remaining' threshold. Financial calculations "
               "are not yet active in this build.")
