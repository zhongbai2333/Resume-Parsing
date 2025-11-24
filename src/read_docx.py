import argparse
import json
import os
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Any, Dict, Iterable, List, Union

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def strip_ns(tag: str) -> str:
    if tag is None:
        return ""
    if tag.startswith("{"):
        return tag.split("}", 1)[1]
    return tag


def collect_text(element: ET.Element) -> str:
    pieces: List[str] = []
    for node in element.iter():
        tag = strip_ns(node.tag)
        if tag in {"t", "delText", "instrText"} and node.text:
            pieces.append(node.text)
        elif tag == "tab":
            pieces.append("\t")
        elif tag in {"br", "cr"}:
            pieces.append("\n")
    return "".join(pieces)


def build_paragraph(paragraph: ET.Element) -> Dict[str, Any]:
    runs: List[str] = []
    for run in paragraph.findall("w:r", NS):
        text = collect_text(run)
        if text:
            runs.append(text)
    text = "".join(runs) if runs else collect_text(paragraph)
    ppr = paragraph.find("w:pPr", NS)
    style = None
    if ppr is not None:
        pstyle = ppr.find("w:pStyle", NS)
        if pstyle is not None:
            style = pstyle.attrib.get(f"{{{NS['w']}}}val")
    return {"type": "paragraph", "style": style, "text": text, "runs": runs}


def build_table(table: ET.Element) -> Dict[str, Any]:
    rows: List[List[Dict[str, Any]]] = []
    for row in table.findall("w:tr", NS):
        cells: List[Dict[str, Any]] = []
        for cell in row.findall("w:tc", NS):
            paragraphs: List[Dict[str, Any]] = []
            for paragraph in cell.findall("w:p", NS):
                paragraphs.append(build_paragraph(paragraph))
            cell_text = "\n".join(
                p["text"] for p in paragraphs if p.get("text")
            ).strip()
            cells.append({"paragraphs": paragraphs, "text": cell_text})
        rows.append(cells)
    max_cols = max((len(r) for r in rows), default=0)
    return {
        "type": "table",
        "row_count": len(rows),
        "column_count": max_cols,
        "rows": rows,
    }


def read_blocks(body: ET.Element) -> List[Dict[str, Any]]:
    blocks: List[Dict[str, Any]] = []
    for child in body:
        tag = strip_ns(child.tag)
        if tag == "p":
            blocks.append(build_paragraph(child))
        elif tag == "tbl":
            blocks.append(build_table(child))
    return blocks


def load_document_xml(path: Union[str, Path]) -> ET.Element:
    if os.path.isdir(path):
        tree = ET.parse(Path(path) / "word" / "document.xml")
    else:
        with zipfile.ZipFile(path) as archive:
            with archive.open("word/document.xml") as doc:
                tree = ET.parse(doc)
    return tree.getroot()


def extract_structure(source: Union[str, Path]) -> Dict[str, Any]:
    document = load_document_xml(source)
    body = document.find("w:body", NS)
    if body is None:
        raise ValueError(f"Could not find body in document: {source}")
    return {"blocks": read_blocks(body)}


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Read the XML structure of a docx without external dependencies."
    )
    parser.add_argument(
        "source",
        help="Path to a .docx file or an extracted docx folder (containing word/document.xml).",
    )
    parser.add_argument(
        "--indent",
        type=int,
        default=2,
        help="JSON indentation level when printing the structure.",
    )
    parser.add_argument(
        "--output",
        "-o",
        help="Optional file path to write the JSON result.",
    )
    args = parser.parse_args()

    structure = extract_structure(args.source)
    payload = json.dumps(structure, ensure_ascii=False, indent=args.indent)

    if args.output:
        Path(args.output).write_text(payload, encoding="utf-8")
        print(f"Structured content written to {args.output}")
    else:
        print(payload)


if __name__ == "__main__":
    main()
