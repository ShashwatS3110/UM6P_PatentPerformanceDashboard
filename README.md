# UM6P Patent Performance Dashboard

**Macro question:** Is UM6P building a patent portfolio that is high-quality, well-protected, globally extended, and generating economic value?

A single-file, self-contained HTML dashboard for tracking UM6P's patent portfolio across 8 tabs: an at-a-glance overview plus scale & momentum, funnel health, portfolio quality, global coverage, commercialization, ventures & partnership risk, and strategic alignment. Each tab leads with one headline indicator and 2–3 supporting ones; many cards open a patent-level drill-down on click.

## Files

- **um6p_dashboard_v3.html** — the dashboard. Open it directly in any browser (Chrome, Edge, Firefox) — no install or server required.
- **UM6P Dashboard Structure - June 16, 2026.xlsx** — the indicator structure, chart logic, and insight rules the dashboard is built from.
- **File 3–8** — source data underlying the dashboard's "Tracked now" indicators:
  - File 3: Master Dashboard + Dataset (Morocco-only)
  - File 4: UM6P Patents Data — Raw (Morocco-only)
  - File 5: Patent Renewal Data — MAScIR (Morocco-only)
  - File 6: Patent Family Data — Espacenet (International)
  - File 7: Citation Data — Orbit (International)
  - File 8: PCT — WIPO (International)

## Status

Every indicator carries a badge: **"Tracked now"** = real figures from the source files above; **"Illustrative"** = placeholder values shown so the full dashboard can be previewed, pending data not yet available (e.g. TTO disclosure/licensing/revenue logs, PATSTAT citation cohorts, per-country family data). Illustrative numbers are flagged in the dashboard rather than hidden or passed off as verified.
