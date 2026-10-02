---
name: bc-order
description: Create draft Microsoft Dynamics 365 Business Central sales orders from customer files, especially Galaxus/Digitec XLSX exports. Use for customer sales orders only; use bc-purchase-order for supplier purchase orders.
---

# BC Sales Order

Create or resume a draft Business Central sales order from a customer order spreadsheet.

## Defaults

- Customer: Digitec Galaxus AG (`80324`)
- Location: `SAM`
- External document number: last long digit group in the filename
- Connector: `/Users/ph/Documents/Silver Spirits Project Codex`
- Credentials: connector `.env`

## Excel Files

Read customer `.xlsx` files through the bundled script or Python workbook libraries. Do not launch Microsoft Excel or require the source workbook to be open. If the source depends on formula results, use cached values or an available headless calculation engine; `openpyxl` does not calculate formulas. If required formula results are unavailable, stop rather than guess or open desktop Excel.

## Workflow

1. Confirm the source XLSX and any overrides (customer, location, external doc no, target sales order, order date). Check the customer's Currency Code in Business Central.
2. Dry-run first.
3. Execute only after validation succeeds.
4. Report the BC sales order number, line count, and totals.

## Currency

- On the sales order header, always explicitly set **Invoice Details > Currency Code** to `EUR` or `CHF`, including when the currency is CHF. Do not leave the field blank to rely on the local-currency default.
- Match the customer's Currency Code in Business Central by default. Use a different currency only when the source order or the user clearly specifies it; validate the order amounts in that currency and state the exception.
- If the source currency conflicts with the customer currency and does not clearly establish an override, resolve the conflict before creating or changing an order.
- The API header field is `currencyCode`. The XLSX script reads the source invoice currency and sends it in that field. Confirm it matches the intended currency in the dry-run and verify the saved order's Currency Code.

## Commands

Dry-run:

```bash
python3 "$HOME/.codex/skills/bc-order/scripts/create_bc_sales_order_from_xlsx.py" "/absolute/path/order.xlsx"
```

Execute:

```bash
python3 "$HOME/.codex/skills/bc-order/scripts/create_bc_sales_order_from_xlsx.py" "/absolute/path/order.xlsx" --execute
```

Target an existing draft sales order:

```bash
python3 "$HOME/.codex/skills/bc-order/scripts/create_bc_sales_order_from_xlsx.py" "/absolute/path/order.xlsx" --sales-order-number B103455 --execute
```

Useful overrides:

- `--customer-number 80324`
- `--location-code SAM`
- `--external-doc-no 196661679`
- `--sales-order-number B103455`
- `--order-date 2026-07-24`
- `--connector-dir "/Users/ph/Documents/Silver Spirits Project Codex"`

## Safety

- Always dry-run before `--execute`.
- Ensure the order header has the intended `currencyCode` before adding lines, and verify the saved Currency Code.
- Prefer `--sales-order-number` when the user names an existing draft order.
- Prefer resume by external document number when no sales order number is given.
- Do not invent BC item numbers. Fail if a supplier item number is missing in BC.
- Keep the order in Draft unless the user asks to release or post it.
