# MGE Project & Billing Command Centre

An executive dashboard for Meraki Global Energy built from the **MGE Project & Invoice Tracker** workbook. It reads two sheets:

- **ASHRAF**, the project register, with one row per job line (2020 onwards).
- **KHAN**, the invoice plan, with monthly planned and actual invoicing per job.

## Files

| File | What it is |
| --- | --- |
| `index.html` | The dashboard. Self-contained HTML, CSS and JavaScript. It loads Chart.js and two Google Fonts from a CDN. |
| `data.json` | Cleaned data produced by `convert.py`. The dashboard reads this file at start-up. |
| `convert.py` | Converts the Excel workbook into `data.json` and records data-quality issues. |
| `README.md` | This file. |

## Quick start

1. Put all four files in one folder.
2. Start a local web server in that folder:
   ```bash
   python -m http.server 8000
   ```
3. Open <http://localhost:8000> in Chrome, Edge, Safari or Firefox.

**Why a server?** Browsers block a page opened by double-clicking (`file://`) from reading `data.json`. If you open `index.html` that way, the dashboard shows a **Choose data.json** button. Pick the file and it works the same.

The machine needs internet access the first time so the charts library and fonts can load. Without it, figures and tables still work but charts show a notice.

## Updating the data

Each time the tracker changes:

```bash
pip install openpyxl            # once
python convert.py "MGE-PROJECT___INVOICE_TRACKER__SUP-UPP_.xlsx"
```

This overwrites `data.json`. Refresh the browser to see the new data.

| Option | Default | Use |
| --- | --- | --- |
| `--asof YYYY-MM-DD` | today | Reporting date. It drives everything that depends on "today": overdue jobs, year-to-date, which plan months count as closed, and dormant clients. |
| `--out FILE` | `data.json` | Output path. |
| `--projects-sheet NAME` | `ASHRAF` | Name of the project register sheet. |
| `--invoices-sheet NAME` | `KHAN` | Name of the invoice plan sheet. |

Example for a board pack dated end of month:

```bash
python convert.py tracker.xlsx --asof 2026-10-31
```

### What the converter tolerates

- **Column order.** Columns are matched by header text, not position, so they can be moved freely.
- **New weekly status columns.** It finds every `WEEK - NN` column and uses the latest two. Labels such as "Week 41 update" follow automatically.
- **New plan months.** It finds every `<MONTH> <YEAR> PLANNED` / `ACTUAL` pair in the invoice sheet. Adding `JANUARY 2027 PLANNED` / `ACTUAL` columns extends the billing runway with no code change.
- **Common entry errors.** It repairs:
  - PO values formatted as dates (Excel shows `19-07-1905` instead of `2027`)
  - progress above 100%
  - `dd-mm-yyyy` dates typed as text

  It flags the rest (see the Data quality view).

Keep these header names unchanged in the tracker:

- **Project register:** `YEAR`, `GROUP`, `COMPANY`, `END-CLIENT`, `JOB CARD STATUS`, `ZOHO SO #`, `DESCRIPTION`, `PROJECT STATUS`, `PROJECT %`, `LPO DATE`, `Customer Order #`, `PO VALUE`, `PS DATE`, `PF DATE`, `ACTUAL START`, `ACTUAL FINISH`, `PROJECT ENGINEER`, `SERVICE ENGINEER`, `FINAL REPORT / DOCUMENTATION`, `WCC SUBMITTED DATE`, `WCC APPROVED DATE`, `SERVICE ENTRY CREATED DATE`, `SERVICE ENTRY NUMBER`, `INVOICE NUMBER`.
- **Invoice plan:** `ZOHO SO #`, `PROJECT DESCRIPTION`, `GROUP`, `Customer Order #`, `PO NUMBER`, `PO VALUE`, `INVOICE MONTH`, `INVOICE AMOUNT`, `TOTAL INVOICE RAISED`, `BALANCE TO INVOICE`, `PAYMENT STATUS`.

Only `DESCRIPTION`, `PROJECT STATUS` and `PO VALUE` are strictly required. Missing optional columns leave blanks.

## Dashboard views

| View | Answers |
| --- | --- |
| **Overview** | Unbilled backlog, earned-not-billed, year-to-date awards vs last year, billing runway, business health score, and generated risks, opportunities and wins. |
| **Order book & growth** | Cumulative awards vs the two previous years, value by year and business line, deal-size profile, and principals/brands. |
| **Billing runway** | Monthly plan vs actual split by confidence, a what-if slider for risk-adjusted billing, balance to invoice by client and month, and the invoice schedule. |
| **Delivery & schedule** | Live-job timeline (Gantt), jobs past planned finish with the latest site note, and age of open jobs since LPO. |
| **Close-out & leakage** | Finished work not yet invoiced, the close-out pipeline (report → WCC → service entry → invoice), cycle times and control gaps. |
| **Clients** | Pareto chart, concentration index, last-year vs this-year comparison, and dormant accounts. |
| **Engineering team** | Live workload and late value by project engineer, site load by service engineer, and a status scorecard. |
| **Project register** | Every job line, sortable and paged, with a detail drawer showing the full lifecycle and week notes. |
| **Data quality** | Every issue `convert.py` found, grouped, with the fix to apply in the tracker. |

**Interactions:**

- **Filters.** The bar at the top filters by year, business line, client, principal, engineer, status and free-text search. Every view follows the same filters.
- **Click to filter.** Clicking a bar in most charts filters the whole dashboard to that item.
- **Job details.** Clicking any job row opens its detail drawer.
- **Export CSV.** Available on Billing, Delivery, Close-out and Register. It exports what is on screen with current filters applied.
- **Light/dark mode.** The setting is remembered in the browser.

## Definitions

| Metric | Definition |
| --- | --- |
| **Unbilled backlog** | PO value of lines at Open PO, In progress or Work completed. |
| **Earned, not billed** | Backlog PO value × reported `PROJECT %`. |
| **Awarded YTD** | Lines whose LPO date falls between 1 January and the as-of date, compared with the same window last year. |
| **Past planned finish** | Open PO or In progress lines with `PF DATE` before the as-of date. |
| **Billing confidence** | Based on the linked job (matched by SO number, then customer PO):<br>• **High:** ≥85% progress, Work completed or Invoiced<br>• **Medium:** 50–84%<br>• **At risk:** below 50%, Open PO, or no matching job.<br>For the current and future months, only the part of the plan not yet invoiced is split this way. |
| **Closed plan month** | A month that started before the as-of month. It shows actual invoicing only. |
| **Overdue to bill** | Open balance whose `INVOICE MONTH` is before the as-of month. |
| **Business health score** | Average of five 0–100 rule-based scores:<br>• order growth<br>• last closed month billing vs plan<br>• share of live value not late<br>• client spread (penalised above 20% for the top client)<br>• data issues per job line.<br>It is a quick read, not an audited measure. |
| **Concentration index** | Herfindahl-Hirschman index of client value shares. Above 2,500 is highly concentrated. |

All values are AED and taken as entered in the tracker. `PO VALUE` is treated as the value of each job line, so several lines can share one customer PO.

## Known limits of the source data

- **Receivables.** The invoice plan has no invoice date, due date or amount received, so overdue receivables and DSO cannot be shown. Adding those three columns would enable them.
- **Schedule baseline.** On most completed rows the planned dates equal the actual dates, which suggests the baseline is overwritten. Keep a separate baseline column if schedule slippage matters.
- **Static snapshot.** The dashboard does not write back to the workbook. Edit the tracker and re-run `convert.py`.

## Hosting

`index.html` and `data.json` are static files, so they can be hosted anywhere that serves static files: SharePoint, an internal web server, Netlify, or an S3 bucket. `data.json` contains commercial data, so host it where only authorised staff can open it.
