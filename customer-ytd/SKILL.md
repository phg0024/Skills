---
name: customer-ytd
description: Build the Silver Spirits Customer YTD Excel analysis from live Business Central data. Use when asked for the Customer YTD report, sales per customer compared to prior years, Swiss/Samnaun/Export customer sales analysis, or the recurring three-week Wednesday customer sales workbook.
---

# Customer YTD

## Purpose

Create the Silver Spirits customer sales workbook from live Business Central posted sales documents.

## Excel Files

Generate and inspect the workbook through the script and Python workbook libraries; do not launch Microsoft Excel or require the output to be open. If formula results need recalculation, use an available headless calculation engine because `openpyxl` does not calculate formulas. If fresh results cannot be obtained headlessly, report that limitation rather than using desktop Excel.

The report compares current year-to-date sales against the same calendar dates in 2025 and 2024, with customer tables and charts split into:

- `Swiss`: sell-to country is Swiss and city is not in the Samnaun area.
- `Samnaun`: sell-to city is `Samnaun`, `Samnaun Dorf`, or `Martina`.
- `Export`: all other customers.

## Workspace

Primary workspace:

```text
/Users/ph/Documents/Silver Spirits Project Codex
```

Bundled skill script:

```text
/Users/ph/.codex/skills/customer-ytd/scripts/build_bc_customer_sales_analysis.py
```

Output folder:

```text
outputs/bc-customer-sales-analysis
```

## Run

From the workspace, run:

```bash
/Users/ph/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 -u /Users/ph/.codex/skills/customer-ytd/scripts/build_bc_customer_sales_analysis.py --as-of YYYY-MM-DD
```

For the current date, omit `--as-of`:

```bash
/Users/ph/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 -u /Users/ph/.codex/skills/customer-ytd/scripts/build_bc_customer_sales_analysis.py
```

The script uses the workspace Business Central credentials from `.env`.

If running from outside the workspace, set:

```bash
CUSTOMER_YTD_WORKSPACE="/Users/ph/Documents/Silver Spirits Project Codex"
```

## Current Report Rules

- Use only posted sales invoices and posted sales credit memos.
- Fetch `salesInvoices + salesInvoiceLines` and `salesCreditMemos + salesCreditMemoLines`.
- Use posting date windows:
  - current YTD: January 1 through the report date
  - 2025 comparison: same calendar dates in 2025
  - 2024 comparison: same calendar dates in 2024
- Credit memos are negative sales and negative quantities.
- Amounts are line `amountExcludingTax`, converted to CHF with Business Central currency exchange rates.
- Include all non-zero line amounts, including non-item lines, because this is a customer revenue report.
- Classify regions from the document header `sellToCountry` and `sellToCity`.

## Expected Workbook

The generated workbook should contain:

- `Dashboard`
- `Region Summary`
- `Customer Detail`
- `Monthly Trend`
- `Swiss Customers`
- `Samnaun Customers`
- `Export Customers`
- `Source Lines`
- `Source Notes`

The dashboard and customer tabs should include native Excel charts. The source lines tab should remain available for auditability.

## Verification

After running the script:

1. Confirm a dated workbook was created in `outputs/bc-customer-sales-analysis`.
2. Inspect the workbook programmatically and verify all expected sheets exist; desktop Excel is not required.
3. Check that `Dashboard`, `Monthly Trend`, and each regional customer tab have charts.
4. Scan the workbook for display/formula errors such as `#REF!`, `#DIV/0!`, `#VALUE!`, `#NAME?`, and `#N/A`.
5. Report the workbook path and the printed regional YTD totals for 2026, 2025, and 2024.

## Automation

The recurring automation should run every three weeks on Wednesday at 09:00 Europe/Zurich time.

Automation prompt:

```text
Use $customer-ytd to generate the Silver Spirits Customer YTD Excel workbook from live Business Central data as of the run date. Verify the expected sheets, charts, source lines, source notes, and absence of formula/display errors, then report the workbook path and regional YTD totals.
```
