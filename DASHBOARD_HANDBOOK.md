# UM6P Patent Performance Dashboard — Handbook

*A one-stop guide to why this dashboard exists, how it is built, how to use it, and how to keep it up to date after handover.*

---

## 1. Why this dashboard

UM6P is filing patents at a record pace, but leadership had no single place to answer a simple strategic question:

> **Is UM6P building a patent portfolio that is high-quality, well-protected, globally extended, and generating economic value?**

Individual spreadsheets could tell you *how many* patents exist, but not whether they are *good*, whether they are *protected where it matters*, or whether they are *turning into value*. This dashboard pulls the scattered data into one view built for **decision-makers**, not analysts — every number is paired with a short, plain-language "so what."

---

## 2. The thinking behind the design

A few principles shaped every choice:

- **Question-led, not data-led.** Each tab answers one question a leader would actually ask ("Are we accelerating?", "Are our patents any good vs peers?", "Are we protected abroad?"). Metrics that don't help a decision were left out.
- **Leading vs supporting indicators.** Each tab has one headline (**leading**) indicator and a few **supporting** ones underneath it.
- **Radical honesty about data.** Every tile is badged **"Tracked now"** (computed from real data) or **"Illustrative"** (a placeholder shown so the structure is visible, awaiting real data). Nothing invented is presented as fact.
- **Two deliberate lenses.** *Volume, themes, coverage, labs and collaboration* use the **full 426-patent inventory**. *Quality and peer benchmarking* use an **internationally-tracked subset (~28 UM6P patents)** — the only patents that appear in international quality databases. These are kept separate on purpose and clearly labelled, because mixing them would mislead.
- **Every headline has a drill-down.** Click any card with a "🔍" hint to see the actual patents behind the number.

---

## 3. What the tabs contain

| Tab | Question it answers | Key content |
|---|---|---|
| **Overview** | Where are we strong / weak vs regional peers? | A radar comparing UM6P to Morocco / Tunisia / South Africa on six quality & impact metrics, plus a summary tile per tab. |
| **Scale & Momentum** | Is filing activity accelerating and healthy? | Filings by year and institution, growth rate + 4-year trend, time-to-grant, disclosures, and the idea-to-revenue funnel. |
| **Labs & Inventors** | Where is IP being created? | Patents per lab (broken down by legal status), most-active inventors, and a top-researchers drill-down. |
| **Research Themes** | Where is activity concentrated, vs strategy? | Theme treemap, momentum over time, portfolio-vs-strategy gap, and citations by theme. |
| **Global Coverage** | How well-protected internationally? | Filing-country treemap, international rate, triadic coverage, survival rate, and average remaining patent life. |
| **Portfolio Quality** | How good are the patents vs peers? | Originality, technology scope, high-impact citations, and science linkage — each vs the selected country average. |
| **Commercialization** | Is IP generating value? | Active licensing rate, deals, revenue (illustrative) and the real portfolio-pruning view. |
| **Industrial Partnership & Ventures** | Are we building industry ties and ventures? | Industry-concentration index, industry–academia collaboration mix, spin-offs (illustrative). |

**Controls:** the "Compare against" toggle (Overview & Quality tabs) switches the peer country. Cards with a "🔍" open a drill-down of the underlying patents; some let you pick a theme, country or lab.

---

## 4. How the data system is built

The dashboard is a **three-layer pipeline** designed so the team never edits code to update numbers — they only replace data files.

```
  Source data files            Build script                 Dashboard
  (Excel / CSV)         ->   build_dashboard_data.ps1  ->   dashboard_data.js  ->  UM6P_Dashboard.html
  (the raw truth)            (reads + computes)             (all numbers)          (what you open)
```

**Layer 1 — Source data files** (kept in the project folder, exact names matter):

| File | Feeds |
|---|---|
| `File 3 - Master Dashboard + Dataset …xlsx` | **The core** — all 426 patents: themes, status, geography, labs, inventors, partners, dates |
| `File 4 …Raw…`, `File 5 …Renewal…` | Grant/renewal dates, time-to-grant, survival rate |
| `File 6 …Escapenet…`, `File 7 …Orbit…`, `File 8 …PCT…` | International families, citations, PCT coverage |
| `morocco_ / tunisia_ / south_africa_assignees_tech_patent_counts.csv` | Real peer benchmarks (quality metrics) |

**Layer 2 — The build script** (`scripts/build_dashboard_data.ps1`) reads those files, computes every metric (counts, rates, breakdowns, drill-down lists, peer averages) and writes them into **`dashboard_data.js`** — a single file holding every number the dashboard shows.

**Layer 3 — The dashboard** (`um6p_dashboard_v3.html`) is a single self-contained page that reads its numbers from `dashboard_data.js`. The refresh step also produces **`UM6P_Dashboard.html`**, a *standalone* copy with the data baked in — that is the file to open or email around.

*Why this design:* the raw data lives in files the team already maintains; the code never needs editing; and one command regenerates everything.

---

## 5. How to VIEW the dashboard

Double-click **`UM6P_Dashboard.html`** (or `View Dashboard.bat`). It opens in any web browser. An internet connection is needed the first time (the charts/fonts load from the web).

---

## 6. How to UPDATE the dashboard (the important part)

When any data file is refreshed (e.g. a new export from OMPIC, or Hicham sends new benchmark data):

1. **Replace the file** in the project folder, keeping the **exact same file name**. (This is the one rule that matters — a renamed file will be ignored.)
2. **Double-click `Refresh Dashboard.bat`.** It checks the files, recomputes every metric, and rebuilds `UM6P_Dashboard.html`. It prints "Done" when finished.
3. **Re-open `UM6P_Dashboard.html`.** The new numbers are in.

That's it — no code, no manual editing. If a file is missing or misnamed, the refresh tool says exactly which one.

**Prerequisites** (already met on the current machine): Windows with PowerShell (built-in), and the **Microsoft Access Database Engine** (present if Microsoft Office/Excel is installed; otherwise install the free "Access Database Engine 2016 Redistributable, 64-bit"). The refresh tool gives this exact message if it's missing.

---

## 7. Reading the badges

- **Tracked now** (green) — computed from real data. Trust it, within the caveats below.
- **Illustrative** (orange) — a placeholder shown so the layout and intent are visible; it turns real automatically once the underlying data file is added.

---

## 8. Known gaps & things to read with care

- **Quality vs peers is a small, international slice** (~28 of 426 UM6P patents) and effectively stops around 2023 — directional, not a full-portfolio verdict.
- **Data gaps in File 3:** ~127 patents have no filing country, ~69 no legal status, ~35 no filing year. Shown honestly; they make the grant rate a *conservative floor*.
- **Grant / survival / time-to-grant** come from the MAScIR renewal registry only, not the full portfolio.
- **Theme = one tag per patent** (UM6P's own classification), so cross-cutting patents are forced into a single bucket.
- **OCP concentration may be over-stated** — framework-reference tags are counted as OCP-linked.
- **Still fully illustrative** (need TTO data): commercialization/licensing/revenue, spin-offs, invention disclosures, output-per-researcher.
- **Benchmarks:** only quality metrics have real peers today (Morocco/Tunisia/South Africa). Extending benchmarks to the other indicators is the top open item, pending data.

*(A fuller gap list and the outstanding data requests are maintained separately for the team.)*

---

## 9. Project files at a glance

| File | What it is |
|---|---|
| `UM6P_Dashboard.html` | **The dashboard** — open this. Regenerated on every refresh. |
| `um6p_dashboard_v3.html` | The source template (edit only if changing design). |
| `dashboard_data.js` | All computed numbers (auto-generated — do not edit by hand). |
| `scripts/build_dashboard_data.ps1` | The engine that turns data files into numbers. |
| `scripts/refresh_dashboard.ps1` | The refresh routine (validate → build → standalone). |
| `Refresh Dashboard.bat` / `View Dashboard.bat` | Double-click helpers. |
| `File 3…` etc. + the three country CSVs | The source data. |

**Repository:** github.com/ShashwatS3110/UM6P_PatentPerformanceDashboard
