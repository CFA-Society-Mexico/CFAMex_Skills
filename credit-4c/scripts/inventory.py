"""Index a source folder of filings and search it.

  python inventory.py index <source_folder> <work_dir> [--only REGEX]
      Extracts PDF text per page to <work_dir>/text/<doc>/p0001.txt, dumps every
      xlsx sheet to <work_dir>/sheets/<doc>/<sheet>.csv, writes inventory.json.
  python inventory.py search <work_dir> "<regex>" [--max 60]
      Prints doc | page | line for every match (case-insensitive).

Read-only on the source folder. Console output is ASCII-only.
"""
from __future__ import annotations

import argparse
import csv
import json
import re
import sys
from dataclasses import asdict, dataclass, field
from pathlib import Path

INVENTORY_JSON = "inventory.json"
TEXT_DIR = "text"
SHEETS_DIR = "sheets"

KIND_PATTERNS: tuple[tuple[str, str], ...] = (
    ("transcript", r"transcript|conference|call"),
    ("annual_report", r"annual|anual|10-k|informe"),
    ("quarterly_report", r"quarter|trimestr|10-q|[1-4]q\d{4}|[1-4]t\d{2}"),
    ("debt_document", r"indenture|prospect|suplemento|credit.?agreement|contrato|covenant|bond|cebur"),
    ("rating_report", r"rating|calificaci|fitch|moody|s&p|hr.?ratings"),
)

SECTION_PATTERNS: dict[str, str] = {
    "balance_sheet": r"estado de situaci|balance sheet|statement of financial position|\[210000\]",
    "income_statement": r"estado de resultados|income statement|statement of (profit|operations)|\[310000\]",
    "cash_flow": r"flujos? de efectivo|cash flows?|\[520000\]",
    "debt_breakdown": r"desglose de cr|\[800001\]|borrowings|deuda financiera|debt maturit",
    "covenants": r"covenant|obligaciones de (hacer|no hacer)|restrictive|razones financieras",
    "leases": r"arrendamiento|lease liabilit",
    "related_party": r"partes relacionadas|related part",
    "audit": r"dictamen|auditor|opinion",
}


@dataclass
class DocEntry:
    path: str
    name: str
    ext: str
    kind: str
    pages: int = 0
    sheets: list[str] = field(default_factory=list)
    sections: dict[str, list[int]] = field(default_factory=dict)
    error: str = ""


def ascii_safe(text: str) -> str:
    return text.encode("ascii", "replace").decode("ascii")


def classify(name: str) -> str:
    lowered = name.lower()
    for kind, pattern in KIND_PATTERNS:
        if re.search(pattern, lowered):
            return kind
    return "other"


def safe_name(name: str) -> str:
    return re.sub(r"[^A-Za-z0-9._-]+", "_", name)


def index_pdf(path: Path, entry: DocEntry, text_root: Path) -> None:
    import pdfplumber  # local import: only needed for PDFs

    doc_dir = text_root / safe_name(path.stem)
    doc_dir.mkdir(parents=True, exist_ok=True)
    with pdfplumber.open(path) as pdf:
        entry.pages = len(pdf.pages)
        for number, page in enumerate(pdf.pages, start=1):
            text = page.extract_text() or ""
            (doc_dir / f"p{number:04d}.txt").write_text(text, encoding="utf-8")
            lowered = text.lower()
            for section, pattern in SECTION_PATTERNS.items():
                if re.search(pattern, lowered):
                    entry.sections.setdefault(section, []).append(number)


def index_xlsx(path: Path, entry: DocEntry, sheet_root: Path) -> None:
    from openpyxl import load_workbook

    doc_dir = sheet_root / safe_name(path.stem)
    doc_dir.mkdir(parents=True, exist_ok=True)
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        for sheet in workbook.worksheets:
            entry.sheets.append(sheet.title)
            target = doc_dir / f"{safe_name(sheet.title)}.csv"
            with target.open("w", newline="", encoding="utf-8") as handle:
                writer = csv.writer(handle)
                for row in sheet.iter_rows(values_only=True):
                    writer.writerow(["" if cell is None else cell for cell in row])
    finally:
        workbook.close()


def build_index(source: Path, work: Path, only: str = "") -> list[DocEntry]:
    if not source.is_dir():
        raise SystemExit(f"[x] source folder not found: {ascii_safe(str(source))}")
    if work.resolve().is_relative_to(source.resolve()):
        raise SystemExit("[x] work dir must not be inside the source folder")
    text_root, sheet_root = work / TEXT_DIR, work / SHEETS_DIR
    entries: list[DocEntry] = []
    for path in sorted(source.rglob("*")):
        if not path.is_file():
            continue
        if only and not re.search(only, str(path.relative_to(source)), re.IGNORECASE):
            continue
        ext = path.suffix.lower()
        if ext not in {".pdf", ".xlsx", ".xlsm", ".csv", ".md", ".txt", ".json", ".yaml", ".yml"}:
            continue
        entry = DocEntry(path=str(path), name=path.name, ext=ext, kind=classify(path.name))
        try:
            if ext == ".pdf":
                index_pdf(path, entry, text_root)
            elif ext in {".xlsx", ".xlsm"}:
                index_xlsx(path, entry, sheet_root)
        except (OSError, ValueError, KeyError) as exc:  # unreadable file: record and continue
            entry.error = f"{type(exc).__name__}: {exc}"
        entries.append(entry)
        print(f"[ok] {ascii_safe(path.name)} kind={entry.kind} pages={entry.pages} sheets={len(entry.sheets)}"
              + (f" ERROR {ascii_safe(entry.error)}" if entry.error else ""))
    work.mkdir(parents=True, exist_ok=True)
    (work / INVENTORY_JSON).write_text(json.dumps([asdict(e) for e in entries], indent=1), encoding="utf-8")
    return entries


def search(work: Path, pattern: str, max_hits: int) -> int:
    regex = re.compile(pattern, re.IGNORECASE)
    hits = 0
    for txt in sorted((work / TEXT_DIR).rglob("p*.txt")):
        for line in txt.read_text(encoding="utf-8").splitlines():
            if regex.search(line):
                print(f"{txt.parent.name} | p{int(txt.stem[1:])} | {ascii_safe(line.strip())[:160]}")
                hits += 1
                if hits >= max_hits:
                    print(f"[..] stopped at {max_hits} hits; narrow the pattern")
                    return hits
    for csv_file in sorted((work / SHEETS_DIR).rglob("*.csv")):
        for row_number, line in enumerate(csv_file.read_text(encoding="utf-8").splitlines(), start=1):
            if regex.search(line):
                print(f"{csv_file.parent.name} | sheet {csv_file.stem} row {row_number} | {ascii_safe(line)[:160]}")
                hits += 1
                if hits >= max_hits:
                    return hits
    print(f"[ok] {hits} hits")
    return hits


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description="Index and search a folder of filings (read-only).")
    sub = parser.add_subparsers(dest="cmd", required=True)
    p_index = sub.add_parser("index")
    p_index.add_argument("source", type=Path)
    p_index.add_argument("work", type=Path)
    p_index.add_argument("--only", default="", help="regex on the relative path; index only matching files")
    p_search = sub.add_parser("search")
    p_search.add_argument("work", type=Path)
    p_search.add_argument("pattern")
    p_search.add_argument("--max", type=int, default=60)
    args = parser.parse_args(argv)
    if args.cmd == "index":
        entries = build_index(args.source, args.work, args.only)
        print(f"[ok] indexed {len(entries)} files -> {ascii_safe(str(args.work / INVENTORY_JSON))}")
        return 0
    search(args.work, args.pattern, args.max)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
