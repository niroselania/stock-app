from __future__ import annotations

import argparse
import os
import re
import sys
from collections import defaultdict
from pathlib import Path
from typing import Iterable

sys.path.insert(0, str(Path(__file__).resolve().parent / "vendor"))

import pandas as pd
from openpyxl import load_workbook


TARGET_COLUMNS = {
    "local": "L",
    "local 1": "L",
    "bari": "B",
    "bari 1": "B",
    "deposito": "D",
    "depósito": "D",
    "exporta": "E",
    "carrito": "C",
    "mayorista": "M",
    "control": "CO",
    "retail": "R",
}


def clean_text(value) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if text.lower() in {"nan", "none"}:
        return ""
    if re.fullmatch(r"\d+\.0", text):
        text = text[:-2]
    return re.sub(r"\s+", " ", text)


def norm_sku(value) -> str:
    text = clean_text(value).upper()
    if not text:
        return ""
    if re.fullmatch(r"\d+", text):
        stripped = text.lstrip("0")
        return stripped or "0"
    return text


def norm_part(value, default: str = "") -> str:
    text = clean_text(value).upper()
    return text or default


def norm_qty(value) -> float:
    if value is None or clean_text(value) == "":
        return 0
    if isinstance(value, (int, float)):
        return value
    text = clean_text(value).replace(".", "").replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return 0


def display_qty(value: float):
    return int(value) if float(value).is_integer() else value


def report_kind(path: Path) -> str | None:
    name = path.stem.lower()
    name = re.sub(r"[_\-]+", " ", name)
    name = re.sub(r"\s+", " ", name).strip()
    for token, column in TARGET_COLUMNS.items():
        if token in name:
            return column
    return None


def split_article(value) -> tuple[str, str]:
    text = clean_text(value).upper()
    if not text:
        return "", ""
    parts = text.split()
    sku = norm_sku(parts[0])
    color = norm_part(parts[1]) if len(parts) > 1 else ""
    return sku, color


def looks_like_header(row: Iterable) -> bool:
    values = [norm_part(v) for v in row]
    return "ARTÍCULO" in values or "ARTICULO" in values or "CANTIDAD" in values


def header_index(row: Iterable) -> dict[str, int]:
    out = {}
    for idx, value in enumerate(row):
        key = norm_part(value)
        if key in {"ARTÍCULO", "ARTICULO"} and "articulo" not in out:
            out["articulo"] = idx
        elif key == "TALLE" and "talle" not in out:
            out["talle"] = idx
        elif key == "CANTIDAD" and "cantidad" not in out:
            out["cantidad"] = idx
    return out


def iter_report_rows(path: Path):
    excel = pd.ExcelFile(path)
    for sheet_name in excel.sheet_names:
        frame = pd.read_excel(path, sheet_name=sheet_name, header=None, dtype=object)
        if frame.empty:
            continue

        first_row = list(frame.iloc[0])
        if looks_like_header(first_row):
            indexes = header_index(first_row)
            data = frame.iloc[1:]
        else:
            indexes = {"articulo": 0, "talle": 1, "cantidad": frame.shape[1] - 1}
            data = frame

        if "articulo" not in indexes or "cantidad" not in indexes:
            continue

        for _, row in data.iterrows():
            sku, color = split_article(row.iloc[indexes["articulo"]])
            talle = norm_part(row.iloc[indexes.get("talle", -1)], "ALL")
            qty = norm_qty(row.iloc[indexes["cantidad"]])
            if not sku or qty == 0:
                continue
            yield (sku, color, talle), qty


def collect_reports(report_paths: Iterable[Path]):
    totals = {column: defaultdict(float) for column in set(TARGET_COLUMNS.values())}
    stats = {}
    for path in report_paths:
        column = report_kind(path)
        if not column:
            stats[path.name] = {"column": None, "rows": 0}
            continue
        rows = 0
        for key, qty in iter_report_rows(path):
            totals[column][key] += qty
            rows += 1
        stats[path.name] = {"column": column, "rows": rows}
    return totals, stats


def stock_key(sheet, row: int, indexes: dict[str, int]) -> tuple[str, str, str]:
    return (
        norm_sku(sheet.cell(row, indexes["SKU"]).value),
        norm_part(sheet.cell(row, indexes["COLOR"]).value),
        norm_part(sheet.cell(row, indexes["TALLE"]).value, "ALL"),
    )


def complete_stock(stock_path: Path, report_paths: Iterable[Path], output_path: Path):
    report_paths = [Path(p) for p in report_paths if Path(p).suffix.lower() in {".xls", ".xlsx"}]
    totals, stats = collect_reports(report_paths)

    workbook = load_workbook(stock_path)
    if "STOCK" not in workbook.sheetnames:
        raise ValueError("La planilla base no tiene una solapa llamada STOCK.")
    sheet = workbook["STOCK"]

    headers = {norm_part(cell.value): cell.column for cell in sheet[1]}
    required = ["SKU", "COLOR", "TALLE", *set(TARGET_COLUMNS.values())]
    missing = [name for name in required if name not in headers]
    if missing:
        raise ValueError(f"Faltan columnas en STOCK: {', '.join(missing)}")

    matched = defaultdict(int)
    for row in range(2, sheet.max_row + 1):
        key = stock_key(sheet, row, headers)
        for column, data in totals.items():
            sheet.cell(row, headers[column]).value = display_qty(data.get(key, 0))
            if key in data:
                matched[column] += 1

    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)

    return {
        "output": str(output_path),
        "reports": stats,
        "matched": dict(matched),
        "loaded": {column: len(data) for column, data in totals.items()},
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Completa STOCK.xlsx con cantidades de reportes.")
    parser.add_argument("--stock", required=True)
    parser.add_argument("--reports-dir")
    parser.add_argument("--reports", nargs="*")
    parser.add_argument("--output", required=True)
    args = parser.parse_args()

    report_paths = []
    if args.reports_dir:
        report_paths.extend(Path(args.reports_dir).glob("*.xls*"))
    if args.reports:
        report_paths.extend(Path(p) for p in args.reports)

    result = complete_stock(Path(args.stock), report_paths, Path(args.output))
    print(result)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
