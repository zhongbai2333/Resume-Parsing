import argparse
import json
import sys
from pathlib import Path
from typing import Any, Dict, Iterable, List, Sequence


def load_structure_from_source(source: str) -> Dict[str, Any]:
    if source == "-":
        raw = sys.stdin.read()
    else:
        raw = Path(source).read_text(encoding="utf-8")
    return json.loads(raw)


def normalized_cell_text(cell: Dict[str, Any]) -> str:
    text = cell.get("text", "")
    if not text:
        paragraphs: Iterable[Dict[str, Any]] = cell.get("paragraphs", [])
        text = "\n".join(p.get("text", "") for p in paragraphs if p.get("text"))
    return text.replace("\n", " ").strip()


def trim_trailing_empty(cells: List[str]) -> List[str]:
    trimmed = cells[:]
    while trimmed and trimmed[-1] == "":
        trimmed.pop()
    return trimmed


def normalize_table(table_block: Dict[str, Any]) -> Dict[str, Any]:
    rows: List[List[str]] = []
    max_cols = 0
    for row in table_block.get("rows", []):
        normalized_row = [normalized_cell_text(cell) for cell in row]
        normalized_row = trim_trailing_empty(normalized_row)
        rows.append(normalized_row)
        max_cols = max(max_cols, len(normalized_row))
    return {
        "row_count": len(rows),
        "column_count": max_cols,
        "cells": rows,
    }


def extract_tables(structure: Dict[str, Any]) -> List[Dict[str, Any]]:
    tables = []
    for block in structure.get("blocks", []):
        if block.get("type") != "table":
            continue
        tables.append(normalize_table(block))
    return tables


def render_table_grid(table: Dict[str, Any]) -> str:
    header = f"{table['row_count']} rows × {table['column_count']} cols"
    lines: List[str] = [header]
    for row in table["cells"]:
        lines.append(" | ".join(cell or "(empty)" for cell in row))
    return "\n".join(lines)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Keep only tables from a read_docx JSON dump and normalize their cells."
    )
    parser.add_argument(
        "source",
        help="Path to read_docx JSON output (use '-' to read from stdin).",
        nargs="?",
        default="-",
    )
    parser.add_argument(
        "--output",
        "-o",
        help="Optional file path to write cleaned tables as JSON.",
    )
    parser.add_argument(
        "--debug",
        action="store_true",
        help="Print each extracted table structure to stdout for debugging.",
    )
    args = parser.parse_args()

    structure = load_structure_from_source(args.source)
    tables = extract_tables(structure)

    if args.debug:
        for index, table in enumerate(tables, start=1):
            print(f"Table {index}: {render_table_grid(table)}")
            print("---")

    payload = json.dumps({"tables": tables}, ensure_ascii=False, indent=2)
    if args.output:
        Path(args.output).write_text(payload, encoding="utf-8")
        print(f"Cleaned tables written to {args.output}")
    else:
        print(payload)


if __name__ == "__main__":
    main()
