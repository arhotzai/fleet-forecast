# Fleet Forecast → Lovable: UI brief + integration guide

This file has two jobs:
1. **Explain how Lovable connects to your existing Python backend** (read Part 0 first).
2. **Give you the exact prompt to paste into Lovable** (Part 1) + the API contract both
   sides agree on (Part 2) + the step-by-step to run it locally and deploy (Part 3).

---

## Part 0 — How this works (read me first)

**Lovable builds the frontend** (React + TypeScript + Tailwind + Vite). It is brilliant at UI,
but it does **not** run Python. Your projection engine (`src/fleet_erp`) is Python and proven —
we keep it. So the shape is:

```
┌────────────────────────┐        REST / JSON over HTTPS        ┌──────────────────────────────┐
│  Lovable React UI      │  ───────────────────────────────►   │  FastAPI backend             │
│  (new repo, Lovable ⇄  │                                      │  (thin layer wrapping         │
│   GitHub two-way sync) │  ◄───────────────────────────────   │   your existing src/fleet_erp)│
└────────────────────────┘        typed JSON responses         └──────────────────────────────┘
        Vercel / Lovable host                                        Railway / Render / Fly.io
```

- **Who builds what:** Lovable builds the UI from the prompt in Part 1. Claude (me) builds the
  FastAPI backend that exposes your engine over the endpoints in Part 2 — say the word and I'll
  add an `api/` module to your `fleet-forecast` repo.
- **Local first:** run the backend locally (`uvicorn` on `http://localhost:8000`) and the Lovable
  app locally pointing at it via an env var. Then deploy both.
- **GitHub + working simultaneously:** in Lovable, open **GitHub → Connect** — it creates a repo
  you own and keeps it in two-way sync. You edit in Lovable *or* push from your machine; both stay
  in step. (It will be a **separate repo** from `fleet-forecast` — the UI and the engine are two
  projects that talk over the API. That's the clean way to work on them in parallel.)
- **Start with mock data:** the prompt tells Lovable to build against the API contract using typed
  mock data and a `VITE_API_BASE_URL` env var. So you get a beautiful, clickable UI on day one;
  flip the env var to the real backend when it's ready — no UI rewrite.

---

## Part 1 — The prompt to paste into Lovable

> Copy everything inside the box into Lovable's prompt. Then use **GitHub → Connect** to create
> the repo.

```
Build "Fleet Forecast" — a premium, dark-themed web app for airline fleet maintenance
forecasting. Think Power BI meets McKinsey: dense, precise, executive-grade, but calm and
uncluttered. This is an aviation ERP module used by airline technical-planning teams.

TECH & ARCHITECTURE
- React + TypeScript + Tailwind + Vite. Use shadcn/ui for primitives and Recharts for charts.
- All data comes from a REST API. Read the base URL from `VITE_API_BASE_URL`
  (default `http://localhost:8000`). Build a typed API client in `src/lib/api.ts`.
- Until the backend is live, use typed MOCK data that exactly matches the API shapes below, so
  every screen is fully clickable. One flag (`VITE_USE_MOCKS`) switches mocks ↔ real API.
- Use TanStack Query for data fetching/caching. Keep all API response types in `src/types.ts`.

DESIGN SYSTEM
- Dark, deep-navy canvas. Palette:
  • background #0b1220, surface #0f1726, card #121c30, border #243049
  • text #e6eaf2, muted #9aa6bd
  • accent (primary) #2d7ff9, accent-2 #1a73c7
  • status: green #2ecc71 (life remaining / live), red #ff6b6b (overdue),
    teal #17b3a3, amber #e0a458, pink #e0557b, blue #5b8def
- Typography: Inter. Big, confident numbers on KPI cards; UPPERCASE micro-labels with letter-spacing.
- Components to nail:
  • KPI card: large value, small label, a colored 3px left border. One card ("Live total fleet")
    has a small pulsing green dot before its label.
  • Pills: a green/red "signed" pill with a glowing dot for remaining life (+ = green, − = red);
    colored "reason" badges (hard link = teal, soft link = amber, LLP/replace = pink, limit = blue).
  • Tables: sticky header, zebra rows, hover highlight, right-aligned tabular numbers, rounded
    container with soft shadow, horizontal scroll inside its own container (page never scrolls x).
  • Charts: centered donut with the total in the hole; stacked bars with per-segment + total
    labels; clean line charts. Distinct colors for NEO / CEO / ATR families.
- Left sidebar navigation with icons, a header band with an airline logo placeholder + "Fleet
  Forecast" wordmark (the word "Forecast" in accent blue), and a prominent file-upload control.

LAYOUT & PAGES (sidebar sections, in this order)
1. Overview / header band + a data-reference banner ("Data reference (snapshot): {date} — the
   projection's starting point"). Below it two rows of KPI cards:
   row 1: Live total fleet (pulsing green dot) · NEO · CEO · ATR aircraft counts
   row 2: NEO spare engines (PW1100) · CEO spare engines (V2500) · ATR spare engines (PW127) · Spare APUs
2. Current Fleet Status — live monitor. Four cascading multi-select filters (Unit kind, Type,
   Event, Tail/unit) where each filter only offers values valid for the ones chosen above it.
   A dense table: Tail · Unit/serial · Kind · Type · Pos · Event · Driver · Since last event ·
   Cycles since SV · Limit · Remaining (green/red signed pill) · % of limit · TSN @ last SV ·
   CSN @ last SV · TSN now · CSN now · Last SV. Sorted closest-to-due first. CSV export button.
3. SV Dashboard — projected shop visits: filters by family/category, a timeline/calendar heat view,
   and a styled events table with colored reason badges.
4. Fleet & Utilisation Stats — charts: fleet composition donut, owner/lessor stacked bars with
   counts, utilisation (annual FH/FC) assumptions table + distribution charts.
5. SV Calendar — month-by-month calendar of due events.
6. Event Log — full projected-event table (unit, event, due date, reason badge).
7. LLP Status — life-limited parts table with a two-tone life bar (consumed ▸ remaining) per part.
8. Aircraft Detail — pick a tail → its airframe checks, engines, APU, propellers and their events.
9. Powerplant — engines & APUs (installed + spare). KPI cards (installed engines, spare engines,
   installed APUs, spare APUs). Table: serial · type · position · state · on-tail · TSN · TSN@LSV ·
   CSN · last-SV date · workscope · #LLPs. Row click → drill-down of that unit's shop visits + LLPs.
10. Edit Data — spreadsheet-like editor of each sheet with "download updated workbook" + "apply".
11. Export — build & download the full month-by-month projection workbook; CSV exports.
12. Guide & Templates — download a sample workbook + a printable 2-page overview.

INTERACTIONS & QUALITY BAR
- Everything is filterable and fast. Empty states are elegant ("Upload your database to begin —
  nothing is pre-loaded"). Loading states use subtle skeletons. Numbers use thousands separators.
- Fully responsive; wide tables scroll inside their own container. Keyboard-accessible.
- No fake branding; use a neutral "AIRLINE" logo placeholder the user can swap.

Build all 12 pages now with mock data wired through the typed API client so the whole app is
navigable, then I will point VITE_API_BASE_URL at the real backend.
```

---

## Part 2 — API contract (both the Lovable client and the FastAPI backend implement this)

Base URL from `VITE_API_BASE_URL`. A session is created by uploading a workbook; its `session_id`
is passed on every subsequent call. All responses are JSON.

| Method & path | Purpose | Returns (shape) |
|---|---|---|
| `POST /api/session/upload` | Upload the `.xlsx`; run ingestion + projection | `{ session_id, snapshot_date, warnings[] }` |
| `GET /api/sample` | Download the sample workbook | xlsx file |
| `GET /api/{sid}/overview` | Global KPI cards | `{ snapshot_date, fleet:{total, neo, ceo, atr}, spares:{neo_eng, ceo_eng, atr_eng, apu} }` |
| `GET /api/{sid}/fleet-status` | Live monitor rows (query: `kind,type,event,tail`) | `{ rows: StatusRow[], filters:{kinds[],types[],events[],tails[]} }` |
| `GET /api/{sid}/projection` | Projected events (event log / SV matrix) | `{ events: ProjEvent[] }` |
| `GET /api/{sid}/calendar` | Month-by-month SV calendar | `{ months: CalendarCell[] }` |
| `GET /api/{sid}/powerplant` | Engines & APUs incl. spares | `{ kpis:{...}, units: PowerUnit[] }` |
| `GET /api/{sid}/llp` | LLP registry status | `{ rows: LlpRow[] }` |
| `GET /api/{sid}/aircraft/{tail}` | Per-aircraft detail | `{ tail, family, events[], engines[], apu, propellers[] }` |
| `GET /api/{sid}/export` | Full projection workbook | xlsx file |

**Core row shapes** (keep these in `src/types.ts`):

```ts
type StatusRow = {
  tail: string; unit: string; kind: string; type: string; pos: string;
  event: string; driver: "FH" | "FC" | "CALENDAR_MONTHS";
  sinceLastEvent: number | null; cyclesSinceSV: number | null;
  limit: number | null; remaining: number | null; pctOfLimit: number | null;
  tsnAtLsv: number | null; csnAtLsv: number | null;
  tsnNow: number | null; csnNow: number | null; lastSv: string | null;
};
type ProjEvent = {
  unit: string; tail: string; kind: string; event: string;
  dueDate: string; reason: string; // reason → badge color
};
type PowerUnit = {
  serial: string; type: string; pos: string; state: "Installed" | "Spare";
  onTail: string | null; tsn: number | null; tsnAtLsv: number | null;
  csn: number | null; lastSvDate: string | null; workscope: string | null; llps: number;
};
type LlpRow = {
  unitId: string; pn: string; sn: string; description: string;
  limitCycles: number; slsv: number; toGo: number;
  projectedSv: string | null; reason: string;
};
```

---

## Part 3 — Step by step

**A. Create the UI (Lovable)**
1. Paste the Part-1 prompt into Lovable → let it build all 12 pages with mock data.
2. In Lovable: **GitHub → Connect** → create a new repo (e.g. `fleet-forecast-ui`). This enables
   two-way sync — you and Lovable now share the repo.
3. Iterate on the look in Lovable. Everything works on mock data already.

**B. Create the backend (I build this in your `fleet-forecast` repo)**
1. I add an `api/` FastAPI app that wraps `src/fleet_erp` and implements every endpoint in Part 2.
2. Run it locally: `uvicorn api.main:app --reload --port 8000`.

**C. Connect them locally**
1. In the Lovable repo set `VITE_API_BASE_URL=http://localhost:8000` and `VITE_USE_MOCKS=false`.
2. Run the UI locally (`npm run dev`). Real data from your engine now flows into the premium UI.

**D. Deploy**
1. Backend → Railway / Render / Fly.io (free tier is fine). Note its public URL.
2. UI → deploy from Lovable (or Vercel). Set `VITE_API_BASE_URL` to the backend's public URL.
3. Add the login/password and CORS allow-list on the backend.

---

### What to do next
Tell me **"build the backend"** and I'll add the `api/` FastAPI layer to your `fleet-forecast`
repo, matching the contract above exactly, so the Lovable UI plugs straight in.
