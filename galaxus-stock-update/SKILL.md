---
name: galaxus-stock-update
description: Update Silver Spirits Galaxusstock.xlsx by deducting item quantities from Business Central Lines*.xlsx exports. Use for Galaxus stock reductions, exact-zero row removal, missing-product reporting, and workbook-preserving verification.
---

# Galaxus Stock Update

Update the Galaxus stock workbook from one or more Business Central line exports without changing unrelated workbook content.

## Excel Files

Read and write the workbooks directly with scripts or Python workbook libraries; do not launch Microsoft Excel or require the workbook to be open. Reload saved workbooks programmatically for verification. Use a headless calculation engine if formulas need recalculation; `openpyxl` does not calculate formulas. If fresh results cannot be obtained headlessly, report that limitation rather than using desktop Excel.

## Defaults

- Stock workbook directory: `/Users/ph/Library/CloudStorage/OneDrive-IZComputer/SilverSpirits/Claude/GalaxusStockUpdate`
- Typical source exports: `Lines*.xlsx` under `/Users/ph/Library/CloudStorage/OneDrive-IZComputer/SilverSpirits/Attachments`
- Prefer a dated output copy under `/Users/ph/Documents/Silver Spirits Project Codex/outputs/galaxusstock_update_YYYYMMDD` unless the user explicitly asks to replace the canonical workbook.

## Rules

1. Read the stock table headers and require `ProductID` and `Inventory`.
2. In each Lines export, include only rows whose `Type` is `Item`; match source `No.` to stock `ProductID` as trimmed strings.
3. Aggregate `Quantity` across all supplied files before changing inventory. Do not count blank, non-item, or zero-quantity lines.
4. Set each matched inventory to `old inventory - aggregated quantity`.
5. Remove rows whose final inventory is exactly zero. Keep negative rows so oversold stock remains visible, and report them explicitly.
6. Report source item numbers that do not exist in Galaxusstock; do not create new stock rows automatically.
7. Preserve headers, table/filter behavior, row order, formulas, styles, widths, and every non-inventory value.

## Safety and verification

- Inspect a mapping and totals before writing. Check for duplicate ProductIDs in the stock workbook and fail if they make matching ambiguous.
- Preserve the source workbook by default and write a separate output copy.
- Reload the saved workbook programmatically and independently verify every deduction, exact-zero removal, unchanged non-inventory cells, table range, row counts, and spreadsheet errors.
- Report files used, products updated, total units deducted, removed products, negative inventories, missing ProductIDs, and the final workbook path.
