"""Validate the raw SKU master and weekly sales CSVs.

Reads the output of generate_data.py, checks every row against the pydantic
models in forecast_common.py plus cross-file checks (duplicate SKUs, sales
rows for unknown SKUs, duplicate SKU/week pairs), and writes:
  - a cleaned SKU master CSV
  - a cleaned sales CSV containing only the rows that passed validation
  - a plain-text report listing every rejected row and why it was rejected

forecast_demand.py is expected to run only on the cleaned CSVs.
"""

import argparse
import csv
import os

from pydantic import ValidationError

from forecast_common import (
    SALES_FIELDS,
    SALES_RAW_PATH,
    SALES_VALIDATED_PATH,
    SKU_FIELDS,
    SKU_MASTER_RAW_PATH,
    SKU_MASTER_VALIDATED_PATH,
    VALIDATION_REPORT_PATH,
    SalesRecord,
    SkuRecord,
)


def read_rows(path: str) -> list[dict]:
    with open(path, newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))


def format_errors(exc: ValidationError) -> str:
    return "; ".join(f"{e['loc'][0]}: {e['msg']}" for e in exc.errors())


def validate_master(rows: list[dict]) -> tuple[list[dict], list[str]]:
    valid_rows: list[dict] = []
    errors: list[str] = []
    seen: set[str] = set()

    for line_no, row in enumerate(rows, start=2):  # header is line 1
        sku = row.get("sku") or f"<row {line_no}>"
        if row.get("sku") in seen:
            errors.append(f"Line {line_no} (sku={sku}): duplicate SKU, dropped")
            continue
        try:
            record = SkuRecord(**row)
        except ValidationError as exc:
            errors.append(f"Line {line_no} (sku={sku}): {format_errors(exc)}")
            continue
        seen.add(record.sku)
        valid_rows.append(row)

    return valid_rows, errors


def validate_sales(rows: list[dict], known_skus: set[str]) -> tuple[list[dict], list[str]]:
    valid_rows: list[dict] = []
    errors: list[str] = []
    seen: set[tuple[str, str]] = set()

    for line_no, row in enumerate(rows, start=2):
        sku = row.get("sku") or f"<row {line_no}>"
        label = f"Line {line_no} (sku={sku}, week={row.get('week_start') or '?'})"

        try:
            record = SalesRecord(**row)
        except ValidationError as exc:
            errors.append(f"{label}: {format_errors(exc)}")
            continue

        if record.sku not in known_skus:
            errors.append(f"{label}: sku not in SKU master, dropped")
            continue

        key = (record.sku, record.week_start.isoformat())
        if key in seen:
            errors.append(f"{label}: duplicate sku/week, dropped")
            continue

        seen.add(key)
        valid_rows.append(row)

    return valid_rows, errors


def write_csv(rows: list[dict], fieldnames: list[str], path: str) -> None:
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def write_report(sections: list[tuple[str, int, int, list[str]]], path: str) -> None:
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        f.write("Demand forecasting data validation report\n")
        f.write("==========================================\n\n")
        for title, total, valid, errors in sections:
            f.write(f"{title}\n")
            f.write(f"{'-' * len(title)}\n")
            f.write(f"Total rows read:   {total}\n")
            f.write(f"Rows passed:       {valid}\n")
            f.write(f"Rows rejected:     {len(errors)}\n")
            if errors:
                f.write("Rejected rows:\n")
                for err in errors:
                    f.write(f"  - {err}\n")
            else:
                f.write("No validation errors found.\n")
            f.write("\n")


def main() -> None:
    parser = argparse.ArgumentParser(description="Validate the raw SKU master and weekly sales CSVs.")
    parser.add_argument("--sales-input", type=str, default=SALES_RAW_PATH, help="Path to the raw sales CSV")
    parser.add_argument("--master-input", type=str, default=SKU_MASTER_RAW_PATH, help="Path to the raw SKU master CSV")
    parser.add_argument(
        "--sales-output", type=str, default=SALES_VALIDATED_PATH, help="Path to write the cleaned sales CSV to"
    )
    parser.add_argument(
        "--master-output", type=str, default=SKU_MASTER_VALIDATED_PATH, help="Path to write the cleaned SKU master to"
    )
    parser.add_argument(
        "--report", type=str, default=VALIDATION_REPORT_PATH, help="Path to write the validation report to"
    )
    args = parser.parse_args()

    master_rows = read_rows(args.master_input)
    valid_master, master_errors = validate_master(master_rows)
    known_skus = {row["sku"] for row in valid_master}

    sales_rows = read_rows(args.sales_input)
    valid_sales, sales_errors = validate_sales(sales_rows, known_skus)

    write_csv(valid_master, SKU_FIELDS, args.master_output)
    write_csv(valid_sales, SALES_FIELDS, args.sales_output)
    write_report(
        [
            ("SKU master", len(master_rows), len(valid_master), master_errors),
            ("Weekly sales", len(sales_rows), len(valid_sales), sales_errors),
        ],
        args.report,
    )

    print(f"SKU master:   {len(master_rows)} rows, {len(valid_master)} passed, {len(master_errors)} rejected")
    print(f"Weekly sales: {len(sales_rows)} rows, {len(valid_sales)} passed, {len(sales_errors)} rejected")
    print(f"Cleaned master -> {args.master_output}")
    print(f"Cleaned sales  -> {args.sales_output}")
    print(f"Report         -> {args.report}")

    if not valid_sales:
        raise SystemExit("No valid sales rows survived validation; nothing to forecast.")


if __name__ == "__main__":
    main()
