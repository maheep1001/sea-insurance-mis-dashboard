# SEA Insurance MIS Dashboard

**A live, single-deployment reporting system for Philippines (PH) and Indonesia (INDO) insurance channel performance — built entirely on Google Apps Script and vanilla JS.**

No servers. No frameworks. No external hosting bill. Just a Google Sheet-backed web app that turns raw policy data into a board-ready Management Information System (MIS) dashboard in real time.

---

## Why this exists

Monthly insurance reporting across two markets usually means stitching together NAV exports, AOP targets, and prior-quarter actuals by hand in Excel — every month, every channel, every currency. This project replaces that spreadsheet shuffle with one dashboard that:

- Pulls **live** policy data straight from the NAV workbooks
- Blends it with **closed-month actuals** and **AOP targets** automatically
- Freezes a period as a locked snapshot for reporting, then **encodes** it back to the source-of-truth Actuals file with one click
- Costs nothing to run beyond a Google account, because the entire stack lives inside Google Apps Script

---

## What it does

| Capability | Details |
|---|---|
| **Dual-country dashboards** | Separate PH (8 channels) and INDO (6 channels) views |
| **Custom date ranges** | Any start/end date across the FY, not just calendar months |
| **Dynamic QTD assembly** | Auto-splits a quarter into "closed" months (from Actuals) + the live current month (from NAV), and sums them |
| **AOP tracking** | Every metric shown against its Annual Operating Plan target, monthly and QTD |
| **Previous-quarter comparisons** | Auto-resolves the correct prior quarter across FY boundaries |
| **Freeze / Unfreeze** | Lock a period's numbers into a snapshot, export to Excel, and keep reporting off live data until you say otherwise |
| **Encode to Actuals** | Writes a frozen snapshot back into the Insurance Actuals workbook — with passthrough-adjusted values for PH's B2B & Affiliates and Agency channels, and safeguards against overwriting existing data |
| **Passthrough modeling (PH)** | Interactive slider lets you choose how pool premium splits between B2B & Affiliates and Agency while holding each channel's blended GM% flat |

### Metrics, per channel

`NOPs` (policy count) · `NP` (Net Premium) · `Rev` (Revenue) · `GM` (Gross Margin) · `Rev%` · `GM%` — all in USD thousands.

---

## Architecture

```
┌─────────────────────────┐
│   index.html (HtmlService) │  ← Dashboard UI + all client logic
│   vanilla JS, no build step │
└─────────────┬───────────┘
              │ google.script.run
              ▼
┌─────────────────────────┐
│   SEA_API.gs (Apps Script)  │  ← doGet() router, data fetch + cache
└─────────────┬───────────┘
              │ reads
              ▼
┌───────────────────────────────────────────────┐
│  4 Google Sheets (source of truth)             │
│  • PH NAV workbook       — live policy data     │
│  • INDO NAV workbook     — live policy data     │
│  • AOP file              — monthly/quarterly targets │
│  • Insurance Actuals     — closed/reported history    │
└───────────────────────────────────────────────┘
```

**Data sourcing rule:** QTD = closed prior months read from the Insurance Actuals file **+** the current reporting month read live from NAV. The current month is never re-read from the closed-actuals file, since that copy may be stale.

**Caching:** Server-side via `CacheService` (10-minute TTL, auto cache-busting on write). Client-side via `localStorage`, used only for freeze-state and frozen snapshots — never as the source for a live read.

---

## Repo contents

| File | Purpose |
|---|---|
| `index.html` | Full dashboard: rendering, date-range logic, freeze/unfreeze state machine, Excel export |
| `SEA_API.gs` | Apps Script backend: `getActuals`, `getActualsFromFile`, `getAOP`, `encodeToActuals`, cache management |
| `PROJECT_SUMMARY.md` | Full technical spec — architecture, phase history, row/column maps, API conventions, testing checklist |
| `HANDOFF_GUIDE.txt` | Deployment and handoff notes |

---

## Getting started

1. Go to [script.google.com](https://script.google.com) and create a new Apps Script project.
2. Replace the default `Code.gs` with the contents of `SEA_API.gs`.
3. Add an HTML file named `index`, and paste in `index.html`.
4. **Deploy → New deployment → Web app**, execute as yourself, access set to your intended audience.
5. Hit `?action=ping` on your deployment URL to confirm all four source files are reachable.
6. Open the deployment URL with no parameters to load the dashboard.

> **Redeploying?** Always version the *existing* deployment (pencil icon → New version) rather than creating a new one — a new deployment changes the endpoint and breaks `google.script.run` calls from any cached client.

Full step-by-step deployment, including optional cache-warming triggers, is in `PROJECT_SUMMARY.md`.

---

## Design principles behind the code

- **Unit discipline**: the Actuals file stores figures in USD thousands; live NAV summaries return raw USD. All scaling happens in one place (`scaleActualsToRawUSD`, `extractPH`/`extractINDO`) so a currency bug can't hide in three different files.
- **Formulas stay formulas**: Rev%/GM% cells are never overwritten with static numbers, even during Encode — they stay derived so future corrections upstream flow through automatically.
- **Snapshot integrity**: Encode always writes from the frozen `localStorage` snapshot taken at freeze time, never a fresh pull — what you reported is what gets written.
- **Guardrails on writes**: `encodeToActuals()` refuses to overwrite non-zero existing data unless explicitly forced.

---

## Status

Production, in active monthly use for PH and INDO reporting cycles. See `PROJECT_SUMMARY.md` for the full phase-by-phase build history and open items.

---

## License

Internal project — no license file included. Add one if you intend to open this up beyond internal use.
