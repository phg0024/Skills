---
name: price-list-update
description: Refresh supplier tabs in Silver Spirits Kalkulation workbooks from assortment or offer files. Use for ProductID-matched price-list updates that preserve names, types, margins, formulas, and color flags.
---

# Price List Update

## Overview

Refresh one supplier tab in a Kalkulation workbook from a supplier price list or assortment workbook. Preserve user-maintained fields by matching ProductID, not by row number.

Use the spreadsheet skill/workflow for workbook handling. Use `/Users/ph/.codex/skills/price-list-update/scripts/update_price_list.py` for the deterministic tab refresh unless the workbook structure requires a custom adjustment.

## Standard Rules

- Update only columns `A`, `B`, `C`, `D`, `E`, `F`, `M`, `N`, `P`, and `Q`.
- Match products by ProductID from target column `M`.
- Preserve target column `B` product name by ProductID. If no old name exists, use the source name.
- Preserve target column `F` product type by ProductID. If no old type exists, use the source type/category.
- Preserve target columns `Y` and `AD` by ProductID. If a product is new, leave `Y` and `AD` empty.
- Preserve yellow fill in target column `B` by ProductID. If a product was yellow before the update, keep it yellow after the update even if its row changes.
- Put new products (ProductIDs not present in the pre-update target tab) in colour blue by applying a solid blue fill (`FF0070C0`) in column `B`.
- If source data is missing for any updated column, leave that target cell empty.
- Clear the updated/preserved columns in old trailing rows after the refreshed source rows.
- Do not overwrite the original workbook unless the user explicitly asks and write permission is available. Prefer writing a dated or descriptive output copy.

## Source Column Mapping

Map supplier columns using header names and common synonyms:

- `M` ProductID: `ProductID`, `Product ID`, `Item code`, `Item Code`, `Art-No Lieferant`, `Supplier ID`
- `A` stock: `Stock`, `Available`, `avlbl cs`, `Btls Available`; if absent, leave empty
- `B` product name: `Description`, `ProductName`, `Product Name`, `Produktname / product name`
- `C` volume percent: `Vol %`, `% Vol`, `Vol`, `Volume`; if absent, parse from the product description
- `D` bottles/case: `Case size`, `btl/Case`, `Btl/case`, `Qty/Case`, `je Krt / per cs`
- `E` size in liters: `size`, `Size`, `Inh / lt`; if absent, parse from the product description
- `F` type/category: `Main Category`, `Category`, `Type`, `Family`
- `N` price per bottle: `Price bottle`, `Price/btl`, `Supplier Price`, `Unit price`; if only unit/case price exists, divide by bottles/case
- `P` currency: `Currency`, `Währung`
- `Q` EAN: `EAN`, `EAN code sales unit`, `barcode`, `Barcode`

For Square-style files, source rows usually start after a header row containing `ProductID`, and `Price bottle` is already per bottle. For ODC/Overseas-style files, `Price/btl` may be formula-derived; load the source workbook with calculated values.

## Workflow

1. Identify the source workbook, target workbook, target supplier tab, and output path.
2. Inspect the source sheets and header row. Prefer the sheet containing `ProductID` or equivalent product-id headers.
3. Run the bundled updater script:

```bash
python3 /Users/ph/.codex/skills/price-list-update/scripts/update_price_list.py --source SOURCE.xlsx --target TARGET.xlsx --tab "Overseas" --output OUTPUT.xlsx
```

Add `--source-sheet`, `--header-row`, or explicit `--*-header` overrides only when auto-detection does not find the right source columns.

4. Verify:
   - refreshed row count matches source product rows
   - B/F/Y/AD preservation has zero errors for matching ProductIDs
   - new products have blank `Y` and `AD`
   - yellow column-B fills are preserved by ProductID
   - new products are marked blue in column B
   - trailing old rows have updated/preserved columns cleared
5. Return a short summary and a Markdown link to the final workbook.

## Important Defaults

- Use the pre-update target workbook as the preservation source. If a previous output exists but the user gives a newer target, preserve from the given target file.
- Preserve formulas outside the updated/preserved columns.
- Product IDs should be compared as trimmed strings; convert integer-like floats such as `123.0` to `123`.
- Empty strings from source should be treated as empty cells.
