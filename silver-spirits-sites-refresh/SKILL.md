---
name: silver-spirits-sites-refresh
description: Refresh Silver Spirits dashboard Site snapshots from canonical Business Central or ABs data. Use for the dormant-products customer tabs, monthly margin snapshot, or their recurring Site refresh actions.
---

# Silver Spirits Sites Refresh

Refresh an existing Silver Spirits dashboard Site snapshot without changing the Site's presentation or access settings.

## Excel Files

Read workbook inputs through the provided scripts or Python workbook libraries; do not launch Microsoft Excel or require the workbook to be open. Use cached formula results or an available headless calculation engine when current calculated values are needed; `openpyxl` does not calculate formulas. If required formula results are unavailable, stop rather than guess or open desktop Excel.

## Workspace

Run from `/Users/ph/Documents/Silver Spirits Project Codex`. The scripts load local `.env` values:

- Builder: `scripts/build_sites_dashboard_data.py`
- Uploader: `scripts/push_site_snapshot.py`

Build the payload completely before uploading it. If source validation, generation, or upload fails, keep the existing live snapshot unchanged and report the exact blocker.

## Monthly margin mode

Use only the canonical workbook:

`/Users/ph/Library/CloudStorage/OneDrive-IZComputer/SilverSpirits/Claude/ABsUpdated/ABs updated.xlsx`

Do not substitute a workbook from `output/` or a dated BC-synced copy unless the user explicitly changes the source.

```bash
python3 scripts/build_sites_dashboard_data.py \
  --dashboard margin \
  --workbook "/Users/ph/Library/CloudStorage/OneDrive-IZComputer/SilverSpirits/Claude/ABsUpdated/ABs updated.xlsx"

python3 scripts/push_site_snapshot.py \
  --dashboard margin \
  --payload margin-3yr-report/app/data/report-data.json
```

Verify the workbook contains the required year tabs, the generated payload has a valid cutoff month and date, and the upload response returns success. Report the workbook, cutoff month, generated date, and response status.

## Dormant customer-tabs mode

```bash
python3 scripts/build_sites_dashboard_data.py \
  --dashboard dormant-customer-tabs \
  --output /tmp/silver-spirits-sites-dormant-customer-tabs.json

python3 scripts/push_site_snapshot.py \
  --dashboard dormant-customer-tabs \
  --payload /tmp/silver-spirits-sites-dormant-customer-tabs.json \
  --env-prefix SILVER_SPIRITS_DORMANT_CUSTOMER_TABS
```

Require:

- every active current-year customer as a tab
- tabs sorted by current-year posted sales-invoice turnover excluding tax, highest first
- at most 100 dormant products per customer, ranked by quantity
- history starting `2025-03-01`, except customer names containing Qoqa use `2024-01-01`
- exclusion of any customer-product with a non-canceled sales-order item line whose quantity remains after `invoicedQuantity`

Verify the JSON shape, tab order, 100-row limit, and successful upload response. Report the as-of and cutoff dates, lookback dates, active customers, opportunities, tab count, open sales-order lines, and excluded open-order products.

## Automation

Update the existing automation IDs rather than creating duplicates:

- `silver-spirits-sites-monthly-margin-refresh`
- `silver-spirits-sites-dormant-products-refresh`
