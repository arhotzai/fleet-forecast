# Fleet Forecast

Aegean fleet **maintenance forecasting** web app — a schema-driven technical projection
engine with a dark, Power BI–style Streamlit dashboard.

Upload a **Fleet Maintenance Status Database** (`.xlsx`), run the projection, and explore
shop-visit calendars, an SV dashboard, LLP life status, and per-aircraft sheets.

> **Confidential.** Fleet and LLP data are **not** stored in this repository — users
> upload their own database file at runtime.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

Then open http://localhost:8501 and upload your database workbook in the sidebar.
(If a `data/FLEET_MAINTENANCE_STATUS_DATABASE.xlsx` is present locally it auto-loads;
that file is gitignored and never committed.)

## Deploy a shareable link (Streamlit Community Cloud)

1. Push this repo to **private** GitHub (already done if you used `gh`).
2. Go to <https://share.streamlit.io> → **New app** → pick this repo, branch `main`,
   main file `app.py`.
3. Open **Advanced settings → Secrets** and paste (see `.streamlit/secrets.toml.example`):
   ```toml
   app_password = "a-strong-password"
   # ANTHROPIC_API_KEY = "sk-ant-..."   # optional, only to enable Ask Claude for everyone
   ```
4. **Deploy.** Share the resulting URL + a database file; recipients upload it and run
   their own projection.

## Security

- `app_password` (Streamlit secret) gates the whole app — share the link **and** the password.
- The Anthropic API key is **never** committed; set it in Streamlit secrets only if you
  want Ask Claude enabled for everyone (it bills your account). Otherwise each user pastes
  their own key on the Ask Claude tab.
- `data/*.xlsx`, `.env.local`, and `.streamlit/secrets.toml` are gitignored.

## Project layout

```
app.py                     Streamlit UI
src/fleet_erp/
  ingestion/loader.py      read + validate the database workbook, merge LLP report
  technical/engine.py      month-by-month projection → events, SV plans, per-aircraft sheets
  export/excel.py          structured workbook export
tools/build_database.py    (maintainer) build the database workbook from source exports
```
