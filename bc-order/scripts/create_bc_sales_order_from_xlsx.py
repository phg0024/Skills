#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import re
import sys
import time
from pathlib import Path
from typing import Any, Callable


DEFAULT_CONNECTOR = Path("/Users/ph/Documents/Silver Spirits Project Codex")


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Create or resume a draft Business Central sales order from a Galaxus/Digitec XLSX order."
    )
    parser.add_argument("xlsx_file", type=Path)
    parser.add_argument("--connector-dir", type=Path, default=DEFAULT_CONNECTOR)
    parser.add_argument("--customer-number", default="80324")
    parser.add_argument("--location-code", default="SAM")
    parser.add_argument("--external-doc-no")
    parser.add_argument("--sales-order-number")
    parser.add_argument("--order-date")
    parser.add_argument("--execute", action="store_true")
    args = parser.parse_args()

    sys.path.insert(0, str(args.connector_dir / "src"))
    from business_central import BusinessCentralClient, BusinessCentralConfig, BusinessCentralError

    try:
        order_rows = read_order_rows(args.xlsx_file)
        external_doc_no = args.external_doc_no or infer_external_doc_no(args.xlsx_file)
        currency = one_value(order_rows, "currency")
        client = BusinessCentralClient(BusinessCentralConfig.from_env(args.connector_dir / ".env"))

        customer = resolve_single(
            client,
            "customers",
            f"number eq '{escape_odata(args.customer_number)}'",
            f"customer number {args.customer_number}",
        )
        location = resolve_single(
            client,
            "locations",
            f"code eq '{escape_odata(args.location_code)}'",
            f"location code {args.location_code}",
        )
        matched_items = match_items(client, order_rows)

        if args.sales_order_number:
            existing_orders = retry(
                "find sales order by number",
                lambda: client.get(
                    client.company_path("salesOrders"),
                    query={
                        "$filter": f"number eq '{escape_odata(args.sales_order_number)}'",
                        "$top": 5,
                    },
                )["value"],
            )
            if not existing_orders:
                raise BusinessCentralError(f"Sales order {args.sales_order_number} was not found.")
            if len(existing_orders) > 1:
                raise BusinessCentralError(
                    f"Sales order number {args.sales_order_number} matched {len(existing_orders)} orders."
                )
        else:
            existing_orders = retry(
                "find existing order",
                lambda: client.get(
                    client.company_path("salesOrders"),
                    query={
                        "$filter": f"externalDocumentNumber eq '{escape_odata(external_doc_no)}'",
                        "$top": 5,
                    },
                )["value"],
            )
            if len(existing_orders) > 1:
                raise BusinessCentralError(
                    f"External document number {external_doc_no} matched {len(existing_orders)} orders."
                )

        summary = {
            "mode": "execute" if args.execute else "dry_run",
            "externalDocumentNumber": external_doc_no,
            "customerNumber": customer.get("number"),
            "customerName": customer.get("displayName"),
            "locationCode": location.get("code"),
            "currencyCode": currency,
            "sourceLineCount": len(order_rows),
            "sourceTotalExcludingTax": round(sum(row["total"] for row in order_rows), 2),
            "matchedItemCount": len(matched_items),
            "existingOrderNumber": existing_orders[0].get("number") if existing_orders else None,
            "targetSalesOrderNumber": args.sales_order_number,
        }

        if not args.execute:
            print(json.dumps({"status": "validated_dry_run", **summary}, indent=2, ensure_ascii=False))
            return 0

        if args.sales_order_number:
            order = existing_orders[0]
            current_external = (order.get("externalDocumentNumber") or "").strip()
            if not current_external and external_doc_no:
                order = retry(
                    "set external document number",
                    lambda: client.patch(
                        client.company_path(f"salesOrders({order['id']})"),
                        {"externalDocumentNumber": external_doc_no},
                        etag=order.get("@odata.etag"),
                    ),
                )
        else:
            order = existing_orders[0] if existing_orders else create_order(
                client,
                customer_number=args.customer_number,
                external_doc_no=external_doc_no,
                currency=currency,
                order_date=args.order_date,
            )
        result = add_missing_lines_and_verify(
            client,
            order,
            order_rows,
            location_id=location["id"],
        )
        print(json.dumps({"status": "created_or_resumed", **summary, **result}, indent=2, ensure_ascii=False))
        return 0
    except Exception as exc:
        parser.exit(status=1, message=f"{exc}\n")


def read_order_rows(path: Path) -> list[dict[str, Any]]:
    try:
        import pandas as pd
    except ImportError as exc:
        raise RuntimeError("pandas is required to read XLSX files. Use the bundled Codex Python runtime.") from exc

    df = pd.read_excel(path, dtype=str).fillna("")
    required = [
        "Supplier item number",
        "Item name",
        "Price per unit",
        "Quantity",
        "Total position",
        "Invoice currency",
    ]
    missing = [column for column in required if column not in df.columns]
    if missing:
        raise RuntimeError(f"Missing required spreadsheet columns: {', '.join(missing)}")

    rows = []
    for _, raw in df.iterrows():
        number = str(raw["Supplier item number"]).strip()
        if not number:
            continue
        rows.append(
            {
                "number": number,
                "description": str(raw["Item name"]).strip(),
                "unitPrice": float(str(raw["Price per unit"]).strip()),
                "quantity": float(str(raw["Quantity"]).strip()),
                "total": float(str(raw["Total position"]).strip()),
                "currency": str(raw["Invoice currency"]).strip(),
            }
        )
    if not rows:
        raise RuntimeError(f"No order rows found in {path}")
    return rows


def infer_external_doc_no(path: Path) -> str:
    matches = re.findall(r"\d{6,}", path.stem)
    if not matches:
        raise RuntimeError("Could not infer external document number; pass --external-doc-no.")
    return matches[-1]


def one_value(rows: list[dict[str, Any]], key: str) -> str:
    values = {str(row[key]).strip() for row in rows if str(row[key]).strip()}
    if len(values) != 1:
        raise RuntimeError(f"Expected exactly one {key}, found {sorted(values)}")
    return next(iter(values))


def resolve_single(client: Any, resource: str, filter_expr: str, label: str) -> dict[str, Any]:
    rows = retry(
        f"resolve {label}",
        lambda: client.get(client.company_path(resource), query={"$filter": filter_expr, "$top": 2})["value"],
    )
    if len(rows) != 1:
        raise RuntimeError(f"Expected one {label}, found {len(rows)}")
    return rows[0]


def match_items(client: Any, rows: list[dict[str, Any]]) -> dict[str, dict[str, Any]]:
    matched = {}
    for row in rows:
        item = resolve_single(
            client,
            "items",
            f"number eq '{escape_odata(row['number'])}'",
            f"item number {row['number']}",
        )
        matched[row["number"]] = item
    return matched


def create_order(
    client: Any,
    *,
    customer_number: str,
    external_doc_no: str,
    currency: str,
    order_date: str | None,
) -> dict[str, Any]:
    payload = {
        "customerNumber": customer_number,
        "externalDocumentNumber": external_doc_no,
        "currencyCode": currency,
    }
    if order_date:
        payload["orderDate"] = order_date
        payload["postingDate"] = order_date
    return retry("create order", lambda: client.post(client.company_path("salesOrders"), payload))


def add_missing_lines_and_verify(
    client: Any,
    order: dict[str, Any],
    rows: list[dict[str, Any]],
    *,
    location_id: str,
) -> dict[str, Any]:
    order_id = order["id"]
    existing_lines = retry(
        "fetch existing lines",
        lambda: client.get(
            client.company_path(f"salesOrders({order_id})/salesOrderLines"),
            query={"$top": 100},
        )["value"],
    )
    existing_numbers = {line.get("lineObjectNumber") for line in existing_lines}
    created = []
    skipped = []

    for row in rows:
        if row["number"] in existing_numbers:
            skipped.append(row["number"])
            continue
        payload = {
            "lineType": "Item",
            "lineObjectNumber": row["number"],
            "quantity": row["quantity"],
            "unitPrice": row["unitPrice"],
            "locationId": location_id,
        }
        retry(
            f"create line {row['number']}",
            lambda payload=payload: client.post(
                client.company_path(f"salesOrders({order_id})/salesOrderLines"),
                payload,
            ),
        )
        created.append(row["number"])
        existing_numbers.add(row["number"])

    verified_order = retry("verify order", lambda: client.get(client.company_path(f"salesOrders({order_id})")))
    verified_lines = retry(
        "verify lines",
        lambda: client.get(
            client.company_path(f"salesOrders({order_id})/salesOrderLines"),
            query={"$top": 100},
        )["value"],
    )
    return {
        "orderNumber": verified_order.get("number"),
        "orderId": order_id,
        "bcStatus": verified_order.get("status"),
        "bcLineCount": len(verified_lines),
        "bcTotalExcludingTax": verified_order.get("totalAmountExcludingTax"),
        "bcTaxAmount": verified_order.get("totalTaxAmount"),
        "bcTotalIncludingTax": verified_order.get("totalAmountIncludingTax"),
        "createdLines": created,
        "skippedExistingLines": skipped,
    }


def retry(label: str, fn: Callable[[], Any], attempts: int = 6) -> Any:
    for attempt in range(1, attempts + 1):
        try:
            return fn()
        except Exception:
            if attempt == attempts:
                raise
            time.sleep(5 * attempt)


def escape_odata(value: str) -> str:
    return str(value).replace("'", "''")


if __name__ == "__main__":
    raise SystemExit(main())
