#!/usr/bin/env python3
import argparse
import re
from copy import copy
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.formula.translate import Translator


REQUIRED_HEADERS = [
    "Description",
    "Unit Volume",
    "Alcohol Percentage Code",
    "Quantity",
    "Unit Price Excl. VAT",
    "Line Discount %",
    "Unit Cost",
]

OPTIONAL_HEADERS = [
    "Document No.",
]

FORMULA_COLUMNS = {7, 10, 14, 15, 18, 19, 20}
SUPPORT_CONSTANT_COLUMNS = {9, 16, 17}
INPUT_COLUMNS = {
    1: "doc",
    2: "desc",
    3: "unit_volume",
    4: "alc",
    5: "qty",
    6: "price",
    8: "discount",
    13: "unit_cost",
}


def natural_key(path):
    # Keep unsuffixed exports first: Lines.xlsx, Lines (1).xlsx, ..., Lines (23).xlsx
    stem = Path(path).stem
    match = re.fullmatch(r"(.*?)(?: \((\d+)\))?", stem)
    base = (match.group(1) if match else stem).lower()
    number = int(match.group(2)) if match and match.group(2) is not None else -1
    return (base, number, Path(path).suffix.lower())


def document_sort_key(value):
    # Sort Document No. values oldest -> newest with natural alphanumeric order.
    if value in (None, ""):
        return (1, ())
    text = str(value).strip()
    if not text:
        return (1, ())
    parts = re.split(r"(\d+)", text)
    key = []
    for part in parts:
        if part == "":
            continue
        if part.isdigit():
            key.append((0, int(part)))
        else:
            key.append((1, part.casefold()))
    return (0, tuple(key))


def normalize_header(value):
    return "" if value is None else str(value).strip().lower()


def normalize_discount(value):
    if value in (None, ""):
        return 0
    if isinstance(value, (int, float)) and value > 1:
        return value / 100
    return value


def source_sheet(workbook):
    if "Lines" in workbook.sheetnames:
        return workbook["Lines"]
    return workbook.active


def header_index(ws):
    headers = {}
    for col in range(1, ws.max_column + 1):
        value = ws.cell(1, col).value
        if value is not None:
            headers[normalize_header(value)] = col
    missing = [header for header in REQUIRED_HEADERS if normalize_header(header) not in headers]
    if missing:
        raise ValueError(f"{ws.title!r} is missing required headers: {', '.join(missing)}")
    return headers


def get_cell(ws, headers, row, header):
    col = headers.get(normalize_header(header))
    if col is None:
        return None
    return ws.cell(row, col).value


def is_blank(value) -> bool:
    if value is None:
        return True
    if isinstance(value, str) and value.strip() == "":
        return True
    return False


def read_lines(paths):
    records = []
    for path in sorted([Path(p) for p in paths], key=natural_key):
        wb = load_workbook(path, data_only=False, read_only=False)
        ws = source_sheet(wb)
        headers = header_index(ws)
        last_doc = None
        for row in range(2, ws.max_row + 1):
            doc = get_cell(ws, headers, row, "Document No.")
            desc = get_cell(ws, headers, row, "Description")
            unit_volume = get_cell(ws, headers, row, "Unit Volume")
            alc = get_cell(ws, headers, row, "Alcohol Percentage Code")
            qty = get_cell(ws, headers, row, "Quantity")
            price = get_cell(ws, headers, row, "Unit Price Excl. VAT")
            discount_raw = get_cell(ws, headers, row, "Line Discount %")
            unit_cost = get_cell(ws, headers, row, "Unit Cost")

            # Keep rows with blank Document No. when other columns have content.
            # Skip only fully empty rows.
            values = (doc, desc, unit_volume, alc, qty, price, discount_raw, unit_cost)
            if all(is_blank(v) for v in values):
                continue

            if isinstance(doc, str):
                doc = doc.strip() or None
            elif is_blank(doc):
                doc = None

            if doc is not None:
                last_doc = doc
                sort_doc = doc
            else:
                # Keep blank Document No. rows with the preceding order in this file.
                sort_doc = last_doc

            records.append(
                {
                    "doc": doc,
                    "sort_doc": sort_doc,
                    "desc": desc,
                    "unit_volume": unit_volume,
                    "alc": alc,
                    "qty": qty,
                    "price": price,
                    "discount": normalize_discount(discount_raw),
                    "unit_cost": unit_cost,
                    "source_order": len(records),
                }
            )

    # Append oldest Document No. first; preserve original line order within each order.
    records.sort(key=lambda record: (document_sort_key(record["sort_doc"]), record["source_order"]))
    return records


def row_is_complete(ws, row):
    core_cols = [1, 2, 5, 6, 13]
    return all(ws.cell(row, col).value not in (None, "") for col in core_cols)


def detect_last_full_row(ws):
    for row in range(ws.max_row, 1, -1):
        if row_is_complete(ws, row):
            return row
    raise ValueError(f"Could not detect a complete data row in sheet {ws.title!r}")


def copy_template_row(ws, template_row, target_row):
    ws.row_dimensions[target_row].height = ws.row_dimensions[template_row].height
    for col in range(1, ws.max_column + 1):
        src = ws.cell(template_row, col)
        dst = ws.cell(target_row, col)

        if src.has_style:
            dst._style = copy(src._style)
        dst.font = copy(src.font)
        dst.fill = copy(src.fill)
        dst.border = copy(src.border)
        dst.alignment = copy(src.alignment)
        dst.protection = copy(src.protection)
        dst.number_format = src.number_format

        dst.value = None
        if col in FORMULA_COLUMNS and isinstance(src.value, str) and src.value.startswith("="):
            dst.value = Translator(src.value, origin=src.coordinate).translate_formula(dst.coordinate)
        elif col in SUPPORT_CONSTANT_COLUMNS:
            dst.value = src.value


def update_workbook(target, sources, output, sheet_name="Abs Updated", last_full_row=None):
    records = read_lines(sources)
    if not records:
        raise ValueError("No source rows were found")

    wb = load_workbook(target, data_only=False)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Target sheet {sheet_name!r} was not found")
    ws = wb[sheet_name]

    template_row = last_full_row or detect_last_full_row(ws)
    start_row = template_row + 1

    for offset, record in enumerate(records):
        row = start_row + offset
        copy_template_row(ws, template_row, row)
        for col, key in INPUT_COLUMNS.items():
            ws.cell(row, col).value = record[key]
        ws.cell(row, 11).value = None
        ws.cell(row, 12).value = None

    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
    except Exception:
        pass

    output = Path(output)
    output.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output)
    return output, len(records), start_row, start_row + len(records) - 1


def main():
    parser = argparse.ArgumentParser(description="Append Business Central Lines exports to ABs updated.xlsx.")
    parser.add_argument("sources", nargs="+", help="Source Lines*.xlsx workbooks")
    parser.add_argument("--target", required=True, help="Target ABs updated.xlsx workbook")
    parser.add_argument("--output", required=True, help="Output workbook path")
    parser.add_argument("--sheet", default="Abs Updated", help="Target sheet name")
    parser.add_argument("--last-full-row", type=int, help="Known last complete row, e.g. 434")
    args = parser.parse_args()

    output, count, start, end = update_workbook(
        target=args.target,
        sources=args.sources,
        output=args.output,
        sheet_name=args.sheet,
        last_full_row=args.last_full_row,
    )
    print(f"Saved {output}")
    print(f"Appended {count} rows at {start}-{end}")


if __name__ == "__main__":
    main()
