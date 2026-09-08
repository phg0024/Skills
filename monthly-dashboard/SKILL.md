---
name: monthly-dashboard
description: Build the bilingual Silver Spirits monthly Business Central Excel dashboard for the previous calendar month. Use for DE/EN monthly dashboards, KPI and chart refreshes, template selection, or the monthly dashboard automation.
---

# Monthly Dashboard

Build and verify the German and English Silver Spirits monthly dashboards from live Business Central data.

## Workspace and script

- Workspace: `/Users/ph/Documents/Silver Spirits Project Codex`
- Generator: `scripts/build_bc_monthly_dashboard_from_template.py`
- Output folder: `output`
- Company: `Silver Spirits KMU`

Run from the workspace with `PYTHONPATH=src` so the Business Central connector resolves correctly.

## Reporting rules

- Report the full calendar month immediately before the run date.
- Compare that month with the same month one year earlier.
- Compare January 1 through report month-end with the same prior-year YTD period.
- Use posted sales invoices, sales credit memos, purchase invoices, and purchase credit memos; treat credit memos as negative.
- Use amounts excluding VAT and convert non-CHF values with Business Central month-end exchange rates.
- Produce separate German and English workbooks with matching data and layout.

## Template selection

1. Prefer the newest validated German dashboard in `output/Silver_Spirits_Dashboard_*_DE.xlsx` from a prior month.
2. Otherwise use the newest valid German dashboard template under `/Users/ph/Library/CloudStorage/OneDrive-IZComputer/SilverSpirits/Attachments`.
3. Require sheets `Dashboard`, `Umsatzanalyse`, and `Einkaufsanalyse`, with charts already present. Do not use a zero-filled or known placeholder workbook as a template.
4. If no valid template exists, stop and report the missing requirement rather than creating an unformatted replacement.

## Run

```bash
cd "/Users/ph/Documents/Silver Spirits Project Codex"
PYTHONPATH=src python3 -u scripts/build_bc_monthly_dashboard_from_template.py \
  --template "/absolute/path/to/validated-DE-template.xlsx" \
  --output-dir output
```

Use `--run-date YYYY-MM-DD` only for a requested historical run. Never use `--dry-run` for a deliverable or scheduled run: it intentionally creates placeholder data.

## Verification

- Confirm both expected DE and EN files were created for the reporting month.
- Confirm `Dashboard`, `Umsatzanalyse`, and `Einkaufsanalyse` exist in both files.
- Confirm chart counts remain at least 3, 3, and 2 on those sheets.
- Scan displayed values and formulas for `#REF!`, `#DIV/0!`, `#VALUE!`, `#NAME?`, and `#N/A`.
- Confirm sales, purchase, document-count, and active-customer KPIs are populated; do not accept an all-zero workbook when source documents exist.
- Check customer and supplier display names do not fall back to bare numeric BC codes.
- Report both output paths, reporting period, month-end FX rates, headline KPIs, and any source caveat.

## Automation

The existing automation id is `silver-spirits-monthly-dashboard`. Update that action instead of creating a duplicate, and invoke this skill explicitly as `$monthly-dashboard` in its prompt.
