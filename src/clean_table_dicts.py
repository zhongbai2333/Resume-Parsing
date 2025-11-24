import argparse
import json
import re
import sys
from pathlib import Path
from typing import Dict, List, Sequence

CHECKED_CHARS = set("☑☒✓✔√■●▣◆▲")
UNCHECKED_CHARS = set("☐□▢◻")
LABEL_KEYWORDS = ["服从调剂", "服从分配", "调剂", "分配"]

FIELD_KEYWORDS: Sequence[tuple[str, Sequence[str]]] = [
    ("姓名", ["姓名"]),
    ("性别", ["性别"]),
    ("出生年月", ["出生年月", "出生"]),
    ("政治面貌", ["政治面貌"]),
    ("所在分院", ["所在分院", "分院"]),
    ("班级", ["班级"]),
    ("学号", ["学号"]),
    ("现（曾） 任职务", ["现任职务", "曾任职", "任职", "现任", "曾任"]),
    ("第一志愿", ["第一志愿"]),
    ("第二志愿", ["第二志愿"]),
    ("联系方式", ["联系方式", "联系电话", "手机", "电话", "手机号"]),
    ("微信", ["微信"]),
    (
        "何时何地曾担任何职务",
        ["何时何地曾担任何职务", "何时何地"]
    ),
    (
        "曾获奖项及获奖时间",
        ["曾获奖项", "曾获", "奖项", "获奖"]
    ),
    (
        "个人优势分析及简要工作设想",
        ["个人优势", "工作设想", "优势分析"]
    ),
    ("服从分配", ["服从分配", "服从调剂"]),
]

ALL_LABEL_KEYWORDS = [
    keyword.replace(" ", "")
    for _, keywords in FIELD_KEYWORDS
    for keyword in keywords
]

LABEL_PREFIX_CHARS = (
    "".join(CHECKED_CHARS)
    + "".join(UNCHECKED_CHARS)
    + ":：-—_/.,()[]（）·•"
)
LABEL_SUFFIX_CHARS = set(":：-—_/.,()[]（）·•")


def load_json(path: str) -> Dict:
    if path == "-":
        raw = sys.stdin.read()
    else:
        raw = Path(path).read_text(encoding="utf-8")
    return json.loads(raw)


def normalize_text(value: str) -> str:
    if not value:
        return ""
    return re.sub(r"[\r\n\s]+", " ", value).strip()


def extract_inline_value(text: str, keyword: str) -> str | None:
    pattern = re.compile(fr"{re.escape(keyword)}\s*[:：]\s*(.+)")
    match = pattern.search(text)
    if match:
        return match.group(1).strip()
    return None


def is_checked_text(text: str) -> bool:
    if not text:
        return False
    t = text.strip()
    is_label_text = any(k in t for k in LABEL_KEYWORDS)

    for ch in t:
        if ch in CHECKED_CHARS:
            return True
        if ch in UNCHECKED_CHARS:
            return False

    if len(t) <= 3:
        if t in ["是", "同意", "接受", "✔", "☑", "√", "✓"]:
            return True
        if t in ["否", "不同意", "不接受"]:
            return False

    if is_label_text:
        return False

    non_alnum = sum(
        1
        for ch in t
        if not ch.isalnum() and not ch.isspace() and "\u4e00" <= ch <= "\u9fff"
    )
    if non_alnum >= 1 and len(t) <= 4:
        if any(ch in t for ch in "■●▣◆"):
            return True
    return False


def interpret_checkbox(text: str) -> str:
    if is_checked_text(text):
        return "是"
    if any(ch in text for ch in UNCHECKED_CHARS):
        return "否"
    return ""


def contains_label_keyword(text: str, allowed_keywords: Sequence[str] | None = None) -> bool:
    normalized = text.replace(" ", "")
    trimmed = normalized.lstrip(LABEL_PREFIX_CHARS)
    if not trimmed:
        return False

    for keyword in ALL_LABEL_KEYWORDS:
        if allowed_keywords and any(keyword == kw.replace(" ", "") for kw in allowed_keywords):
            continue
        if not trimmed.startswith(keyword):
            continue

        remainder = trimmed[len(keyword):]
        if not remainder:
            return True
        if all(ch in LABEL_SUFFIX_CHARS for ch in remainder):
            return True

    return False


def infer_value_from_adjacent(
    rows: List[List[str]], row_idx: int, col_idx: int, allowed_keywords: Sequence[str] | None = None
) -> str:
    row = rows[row_idx]
    # try to the right
    for cc in range(col_idx + 1, len(row)):
        candidate = normalize_text(row[cc])
        if candidate and not contains_label_keyword(candidate, allowed_keywords):
            return candidate
    # try below in same column
    for rr in range(row_idx + 1, len(rows)):
        if col_idx < len(rows[rr]):
            candidate = normalize_text(rows[rr][col_idx])
            if candidate and not contains_label_keyword(candidate, allowed_keywords):
                return candidate
    return ""


def match_field(value: str, keywords: Sequence[str]) -> str:
    normalized = value.replace(" ", "")
    for keyword in keywords:
        if keyword.replace(" ", "") in normalized:
            return keyword
    return ""


def clean_table(table: Dict) -> Dict[str, str]:
    rows = table.get("cells", [])
    normalized_rows = [[normalize_text(cell) for cell in row] for row in rows]
    cleaned = {field: "" for field, _ in FIELD_KEYWORDS}

    for r_idx, row in enumerate(normalized_rows):
        for c_idx, text in enumerate(row):
            if not text:
                continue
            for field, keywords in FIELD_KEYWORDS:
                if cleaned[field]:
                    continue
                matched_keyword = match_field(text, keywords)
                if not matched_keyword:
                    continue
                value = extract_inline_value(text, matched_keyword)
                if not value:
                    value = infer_value_from_adjacent(
                        normalized_rows, r_idx, c_idx, keywords
                    )
                if field == "服从分配":
                    checkbox = interpret_checkbox(text)
                    if checkbox:
                        cleaned[field] = checkbox
                        break
                    checkbox = infer_value_from_adjacent(normalized_rows, r_idx, c_idx)
                    if checkbox:
                        cleaned[field] = interpret_checkbox(checkbox) or checkbox
                        break
                    if value:
                        cleaned[field] = interpret_checkbox(value) or value
                        break
                if value:
                    cleaned[field] = value
                    break
    return cleaned


def clean_tables(data: Dict) -> List[Dict[str, str]]:
    tables = data.get("tables") or []
    return [clean_table(table) for table in tables]


def render_debug(cleaned: List[Dict[str, str]]) -> str:
    lines = []
    for idx, entry in enumerate(cleaned, start=1):
        lines.append(f"Entry {idx}")
        for field in entry:
            lines.append(f"  {field}: {entry[field]}")
        lines.append("---")
    return "\n".join(lines)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Clean the tables produced by extract_tables.py into labeled dictionaries."
    )
    parser.add_argument("source", help="Path to extract_tables JSON output or '-' for stdin.")
    parser.add_argument(
        "--output",
        "-o",
        help="Optional path to write the cleaned dictionaries as JSON.",
    )
    parser.add_argument(
        "--debug",
        action="store_true",
        help="Print the extracted dictionaries after mapping.",
    )
    args = parser.parse_args()

    data = load_json(args.source)
    cleaned = clean_tables(data)

    if args.debug:
        print(render_debug(cleaned))

    payload = json.dumps({"entries": cleaned}, ensure_ascii=False, indent=2)
    if args.output:
        Path(args.output).write_text(payload, encoding="utf-8")
        print(f"Cleaned dictionaries written to {args.output}")
    else:
        print(payload)


if __name__ == "__main__":
    main()
