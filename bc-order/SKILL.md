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

## Workflow

1. Confirm the source XLSX and any overrides (customer, location, external doc no, target sales order, order date).
2. Dry-run first.
3. Execute only after validation succeeds.
4. Report the BC sales order number, line count, and totals.

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
- Prefer `--sales-order-number` when the user names an existing draft order.
- Prefer resume by external document number when no sales order number is given.
- Do not invent BC item numbers. Fail if a supplier item number is missing in BC.
- Keep the order in Draft unless the user asks to release or post it.
