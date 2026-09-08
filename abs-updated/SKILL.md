---
name: abs-updated
description: Update Silver Spirits "ABs updated.xlsx" from live Business Central sales orders or Lines*.xlsx exports. Use for duplicate-safe ABs imports, alcohol-code lookup, and preserved formulas and formatting.
---

# ABs Updated

## Overview

Append new Business Central sales-order lines to the Silver Spirits `ABs updated.xlsx` workbook.

Use the spreadsheet workflow for workbook handling. Prefer these project scripts:

- `/Users/ph/.codex/skills/abs-updated/scripts/update_abs.py` for downloaded `Lines*.xlsx` exports.
- `scripts/sync_abs_from_business_central.py` for live Business Central sales orders.

## Canonical Workbook Location

Use this OneDrive directory as the source/base location:

`/Users/ph/Library/CloudStorage/OneDrive-IZComputer/SilverSpirits/Claude/ABsUpdated`

- Use `/Users/ph/Library/CloudStorage/OneDrive-IZComputer/SilverSpirits/Claude/ABsUpdated/ABs updated.xlsx` as the canonical workbook on the first run.
- Ignore the OneDrive lock file `~$ABs updated.xlsx`.
- After a successful run creates a dated `ABs updated - YYYY-MM-DD - BC synced.xlsx` copy in this same directory, use the newest validated dated copy there as the next base so orders are not imported twice.
- Never use a workbook in the project `output/` folder as the base unless the user explicitly asks for that.
- Keep the canonical/source workbook unchanged by default; write the validated dated output copy in the canonical OneDrive directory.

## Standard Rules

- Use an existing `ABs updated.xlsx`-style workbook and the `Abs Updated` sheet unless the user specifies otherwise.
- Add lines after the last complete table line, not after prepared blank formula rows.
- Preserve formatting, colours, row height, number formats, formulas, and workbook structure from the last complete row.
- Fill only these columns:
  - `A` Document No.
  - `B` Description
  - `C` Unit Volume
  - `D` Alcohol Percentage Code
  - `E` Quantity
  - `F` Unit Price Excl. VAT
  - `H` Line Discount %
  - `M` Unit Cost
- Leave column `K` unchanged unless the user explicitly provides customer names.
- Leave column `L` unchanged unless the user explicitly asks to update Location Code.
- Convert source discount percentages such as `5` to `0.05`, matching the workbook formulas.
- Sort appended orders oldest to newest by `Document No.` and preserve line order within each order.
- For live sync, skip exact document numbers already in the workbook. Also use the highest existing numeric order ID as a cutoff: do not import older open orders that are absent from the workbook.
- Never overwrite the original workbook by default. Write a descriptive output copy.

## Source Column Mapping

For downloaded files, read the `Lines` sheet or active sheet using these headers:

- `Document No.` -> target `A`
- `Description` -> target `B`
- `Unit Volume` -> target `C`
- `Alcohol Percentage Code` -> target `D`
- `Quantity` -> target `E`
- `Unit Price Excl. VAT` -> target `F`
- `Line Discount %` -> target `H`
- `Unit Cost` -> target `M`

Allow blank `Document No.` cells when the preceding order is clear; inherit the preceding order number from the same source file. Skip rows only when they have no meaningful content. Keep shipment-note rows with a description even when quantity, price, volume, and cost are zero.

For live Business Central sync:

- Read order headers from the standard `salesOrders` endpoint and use `number` as the document/order ID and `id` as the source order identifier.
- Read lines from the OData resource `Sales_Order_Line_Excel`, matching `Document_No` to the order `number` and ordering by `Line_No`.
- Map `Description`, `Unit_Volume`, `Quantity`, `Unit_Price`, `Line_Discount_Percent`, and `Unit_Cost` into the target fields above. Match the browser alcohol lookup by document number plus item `No`.
- The standard endpoint may not expose `Alcohol_Percentage_Code`. Do not guess it and do not write an incomplete workbook without explicit approval.

## Downloaded-Export Workflow

1. Identify the target workbook, all `Lines*.xlsx` sources, and a separate output path.
2. Run the deterministic updater:

```bash
python3 /Users/ph/.codex/skills/abs-updated/scripts/update_abs.py \
  --target TARGET.xlsx \
  --output OUTPUT.xlsx \
  SOURCE_LINES.xlsx "SOURCE_LINES (1).xlsx"
```

Use `--sheet "Abs Updated"` or `--last-full-row N` when needed. Inspect the files before running if the target sheet or last complete row is ambiguous.

## Live Business Central Workflow

Use the live sync script from `/Users/ph/Documents/Silver Spirits Project Codex`, using the canonical OneDrive workbook as the target and writing the dated copy beside it:

```bash
python3 scripts/sync_abs_from_business_central.py \
  --target "/Users/ph/Library/CloudStorage/OneDrive-IZComputer/SilverSpirits/Claude/ABsUpdated/ABs updated.xlsx" \
  --output "/Users/ph/Library/CloudStorage/OneDrive-IZComputer/SilverSpirits/Claude/ABsUpdated/ABs updated - YYYY-MM-DD - BC synced.xlsx" \
  --alcohol-json "tmp/abs_browser_alcohol.json" \
  --dry-run
```

Replace `YYYY-MM-DD` with the run date. Run the dry-run first. The script reports the existing document count, latest existing order ID, new orders, new line rows, and missing alcohol values. Remove `--dry-run` only after the plan is correct.

### Browser Alcohol Percentage Lookup

When the standard Business Central endpoint does not return the alcohol code:

1. Use the authenticated Business Central sales-order page in the browser.
2. Open each new order identified by the sync dry-run, searching by its order/document number.
3. For every product line, read `No.` and `Alcohol Percentage Code`. Ignore blank trailing rows and do not infer values.
4. Save the results as JSON keyed by document number and item number, for example:

```json
{
  "B103530": {
    "203354": 40,
    "201136": 53.7
  }
}
```

5. Rerun the dry-run with `--alcohol-json`. Continue only when no required alcohol values are missing.

The browser lookup requires an authenticated interactive session. A scheduled run can be fully unattended only if the alcohol field is exposed through Business Central OData/a custom API or the JSON lookup has already been prepared. If the field is unavailable, stop and report the missing values rather than writing partial rows. Use `--allow-missing-alcohol` only when the user explicitly approves incomplete alcohol data.

For a targeted retry, pass `--only-order B103530` one or more times. Keep the original workbook untouched while collecting or reviewing browser values.

## Verification

Before delivering the workbook:

- Confirm new rows start at the last complete row plus one and the row count matches the source lines.
- Confirm live sync imported only order IDs newer than the workbook cutoff and did not duplicate an existing order.
- Confirm columns `A`, `B`, `C`, `D`, `E`, `F`, `H`, and `M` contain the intended values; confirm appended `K` and `L` remain unchanged/blank as required.
- Confirm calculated-column formulas translate to each new row.
- Confirm appended rows match the template row's styles, number formats, fills, borders, alignment, and row height.
- Run a second dry-run against the generated workbook; it should report zero new orders and zero new line rows.
- Confirm the original target workbook remains unchanged and return a Markdown link to the output copy.

When the target's latest order is `B103529`, for example, a run that finds `B103530` and `B103531` should import those orders while ignoring older missing open orders. This cutoff behavior is intentional.

## Notes

- For exports, read unsuffixed `Lines.xlsx` before `Lines (1).xlsx`, `Lines (2).xlsx`, and so on, then sort all collected rows oldest to newest.
- The target workbook may contain prepared blank formula rows below the last full line; treat those as available rows to fill.
- If files are outside the workspace, request read access for sources and read/write or copy-output access for the target location.
