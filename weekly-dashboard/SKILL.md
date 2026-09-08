---
name: weekly-dashboard
description: Build or maintain the Silver Spirits weekly Business Central Excel dashboard. Use when asked for the Weekly dashboard, weekly BC dashboard, Monday dashboard automation, or changes to the dashboard KPIs, gross margin rules, workbook output, or scheduled report run.
---

# Weekly Dashboard

## Purpose

Create the Silver Spirits Business Central Excel dashboard from live BC data and keep the Monday automation aligned with the current business rules.

## Workspace

Primary workspace:

```text
/Users/ph/Documents/Silver Spirits Project Codex
```

Generator script:

```text
scripts/build_bc_weekly_dashboard.py
```

Output folder:

```text
outputs/bc-weekly-dashboard
```

## Run

From the workspace:

```bash
PYTHONPATH=src python3 -u scripts/build_bc_weekly_dashboard.py
```

For a fixed report date:

```bash
PYTHONPATH=src python3 -u scripts/build_bc_weekly_dashboard.py --as-of YYYY-MM-DD
```

The script uses the existing Business Central credentials in the workspace environment.

## Current Dashboard Rules

- Sales use posted sales invoices and posted sales credit memos.
- Purchases use posted purchase invoices and posted purchase credit memos.
- Credit memos are treated as negative values and quantities.
- Amounts are excluding VAT and converted to CHF using BC currency exchange rates.
- Weeks start on Monday.
- Sales orders are orders created during the current week.
- Item stockout view shows items sold in the past 3 weeks whose current inventory is `<= 0`.
- Gross margin is simple: `sales CHF - (unit cost CHF * quantity)`.
- Gross margin only includes export customers plus Qoqa and House of Single Malt.
- Swiss customers are excluded from gross margin unless the customer name contains Qoqa or House of Single Malt.
- Do not add Swiss alcohol or bottle tax adjustments to margin.

## Expected Workbook

The generated workbook should contain these sheets:

- `Dashboard`
- `Sales by Customer`
- `Sales by Item`
- `Stockouts`
- `Purchases`
- `Gross Margin`
- `Margin Scope`
- `Source Notes`

## Verification

After running the script:

1. Confirm a dated workbook was created in `outputs/bc-weekly-dashboard`.
2. Open or inspect the workbook and verify all expected sheets exist.
3. Check `Dashboard` has KPI values and native Excel charts.
4. Check `Gross Margin` only contains the simple margin scope.
5. Check `Margin Scope` documents included and excluded customers.
6. Report the workbook path and the printed sales, purchase, and sales-order totals.

## Automation

The existing automation id is:

```text
weekly-bc-dashboard
```

When updating its schedule, prefer updating this automation instead of creating a duplicate.
