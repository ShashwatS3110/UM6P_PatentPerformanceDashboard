# UM6P Patent Performance Dashboard

**The question it answers:** Is UM6P building a patent portfolio that is high-quality, well-protected internationally, and generating economic value?

An interactive dashboard over UM6P's 426-patent portfolio, organised into eight tabs — Overview, Scale & Momentum, Labs & Inventors, Research Themes, Global Coverage, Portfolio Quality, Commercialization, and Industrial Partnership & Ventures. Each tab leads with one headline indicator and a few supporting ones; most cards open a patent-level drill-down on click.

## 👉 New here? Start with the docs

- **[DASHBOARD_OVERVIEW.md](DASHBOARD_OVERVIEW.md)** — 2-page plain-language guide: what the dashboard says and the key points to focus on. **Read this first.**
- **[DASHBOARD_HANDBOOK.md](DASHBOARD_HANDBOOK.md)** — the full handbook: design thinking, data pipeline, how to use, how to update, and known gaps.

## ▶️ To VIEW the dashboard

Open **`UM6P_Dashboard.html`** in any web browser (or double-click **`View Dashboard.bat`**). It is fully self-contained — no install, no server. (Needs internet the first time to load the charts.)

## 🔄 To UPDATE it with new data

1. Replace the relevant data file in this folder, **keeping the exact same file name**.
2. Double-click **`Refresh Dashboard.bat`** — it re-reads the data, recomputes every metric, and rebuilds `UM6P_Dashboard.html`.
3. Re-open `UM6P_Dashboard.html`.

No code editing required. See the handbook for prerequisites and troubleshooting.

## How it's built

`Source data files → scripts/build_dashboard_data.ps1 → dashboard_data.js → UM6P_Dashboard.html`

- **Source data:** `File 3–8` (patent inventory, renewals, families, citations, PCT) + `morocco/tunisia/south_africa_assignees_tech_patent_counts.csv` (peer benchmarks).
- **`um6p_dashboard_v3.html`** — the design template (edit only to change the dashboard itself).
- **`dashboard_data.js`** — every computed number (auto-generated; do not edit by hand).

## Tracked vs. Illustrative

Every tile is badged **"Tracked now"** (real, computed from the data) or **"Illustrative"** (a placeholder shown so the layout is complete, which fills in automatically once its data file is added). Nothing invented is presented as verified.

Note: *volume, themes and coverage* use the full 426-patent inventory; *quality and peer benchmarking* use an internationally-tracked subset (~28 patents) and are directional — the dashboard labels this everywhere.
