---
name: stock-update
description: Refresh Silver Spirits Kalkulation "Lager (2)" from a Business Central Stock List.xlsx export. Use for product-ID matching, preserved calculation fields and colors, and blue flags for unmatched products.
---

# Stock Update

## Excel Files

Read and write workbooks directly with the bundled script or Python workbook libraries; do not launch Microsoft Excel or require either workbook to be open. Verify saved files programmatically. Use a headless calculation engine when formulas need recalculation; `openpyxl` does not calculate formulas. If fresh results cannot be obtained headlessly, report that limitation rather than using desktop Excel.

## Workflow

Use this skill for the recurring Silver Spirits stock-list update:

1. Identify the Kalkulation workbook and the Stock List workbook. Default paths are usually under `/Users/ph/Library/CloudStorage/OneDrive-IZComputer/SilverSpirits/Attachments/`.
2. Check whether a lock/temp file such as `~$Kalkulation.xlsx` exists. If present, write a separate output copy instead of overwriting the original.
3. Run `scripts/update_lager2.py` with explicit input paths.
4. Verify the first two updated rows, the last updated row, formula continuation in calculation columns, and the count of blue unmatched products with targeted checks. Avoid importing or rescanning the entire workbook for this fixed-layout refresh. Render only the affected `Lager (2)` view from a temporary copy with a headless office export.
5. Return a Markdown link to the finished `.xlsx`.

## Column Mapping

Update `Lager (2)` from `Stock List` as follows:

| Lager (2) column | Source |
| --- | --- |
| A `SAM` | `SAM - Quantity Available` |
| B `BOZ` | `BOZ quantity available` |
| C `Produktname / product name` | `Product name` |
| D `% Vol` | `Alcohol %` |
| E | `Bottles per case` |
| F `Inh / lt` | `Size` |
| N `Art-No Lieferant` | `Article ID` |
| O `EK` | `Internal price` |
| P `Rabatt` | `0` because Stock List has no discount column |
| R `EAN` | `GTIN` / `EAN` |

Match by product ID: `Stock List!Article ID` (or `Product ID` / `Product UD` in exports with alternate headers) to `Lager!N`.

From matched `Lager` rows, copy:

| Lager (2) column | Matched source |
| --- | --- |
| G `Type` | `Lager!G` |
| Z | `Lager!Z` |
| AE | `Lager!AE` |
| AK | `Lager!AK` |
| C fill color | `Lager!C` fill color |

If no matching product ID exists in `Lager`, leave G, Z, AE, and AK empty and set the product-name cell in column C to solid blue (`FF0070C0`). Do not color the quantity cell in column B.

## Script

Prefer the bundled script:

```bash
python3 /Users/ph/.codex/skills/stock-update/scripts/update_lager2.py \
  --kalkulation "/path/to/Kalkulation.xlsx" \
  --stock-list "/path/to/Stock List.xlsx"
```

Optional:

```bash
--output "/path/to/Kalkulation - Lager 2 updated.xlsx"
--overwrite
```

Use the bundled Python runtime when available. The script uses `openpyxl`, preserves workbook structure as much as practical, extends formulas/styles for extra stock rows from the last non-empty template row in `Lager (2)`, and deletes surplus old rows when the stock list is shorter.
