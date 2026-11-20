"""Build capture.csv from cited extract files, or validate a capture.csv.

  python capture_from_extracts.py <inputs_dir> <out_dir> --years 2021-2025 [--scale 1.0]
      Reads every extract_*.json in <inputs_dir> (rows with line_item, period like
      4Q2025, value, basis, statement, source_doc, source_ref), selects one cited
      fiscal-year value per registry line item, writes <out_dir>/capture.csv and
      <out_dir>/restatements.csv, then validates.
  python capture_from_extracts.py --validate <out_dir>/capture.csv
      PASS/FAIL validation; exit 1 on FAIL.

Selection rules (why: a fiscal-year credit ratio needs a full-year flow and a
year-end balance, and the most recent filing carries any restatement):
  * balance items: basis point_in_time at 4Q<year>
  * flows: basis year_to_date at 4Q<year>; else the sum of four 'quarter' values,
    tagged derived and citing the four sources
  * several filings report the same fact: keep the latest filing; log differences
"""
from __future__ import annotations

import argparse
import csv
import json
import math
import re
import sys
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path

from registry import (
    BY_KEY,
    CAPTURE_COLUMNS,
    CAPTURE_CSV,
    ITEMS,
    RESTATEMENTS_CSV,
    VALID_TAGS,
    Sign,
    Statement,
    fy_period,
    out_path,
    period_year,
)

# research-analyst extract key -> registry key, in priority order per registry key.
SOURCE_KEYS: dict[str, tuple[str, ...]] = {
    "revenue": ("revenue_total", "revenue_net_sales"),
    "operating_profit": ("operating_profit",),
    "depreciation_amortization": ("depreciation_amortization",),
    "ebitda_reported": ("ebitda",),
    "interest_expense": ("financial_expense",),
    "interest_income": ("financial_income",),
    "pretax_profit": ("pretax_profit",),
    "income_tax_expense": ("income_tax_expense",),
    "net_income": ("net_income",),
    "cash_and_equivalents": ("cash_and_equivalents",),
    "receivables": ("trade_and_other_receivables",),
    "inventories": ("inventories",),
    "total_current_assets": ("total_current_assets",),
    "ppe_net": ("ppe_net",),
    "investment_property": ("investment_property",),
    "rou_assets": ("rou_assets",),
    "goodwill": ("goodwill",),
    "intangibles_net": ("intangibles_net",),
    "total_assets": ("total_assets",),
    "payables": ("trade_and_other_payables",),
    "total_current_liabilities": ("total_current_liabilities",),
    "lease_liabilities_current": ("lease_liabilities_current",),
    "lease_liabilities_noncurrent": ("lease_liabilities_noncurrent",),
    "total_equity": ("total_equity",),
    "total_liabilities_and_equity": ("total_liabilities_and_equity",),
    "cfo": ("cfo",),
    "capex": ("capex",),
    "cfi": ("cfi",),
    "cff": ("cff",),
    "interest_paid": ("cf_interest_paid",),
    "lease_interest_paid": ("cf_lease_interest",),
    "income_tax_paid": ("cf_income_tax_paid",),
    "lease_principal_paid": ("lease_payments",),
    "dividends_paid": ("dividends_paid",),
    "share_buybacks": ("share_buybacks",),
    "cash_at_period_start": ("cash_at_period_start",),
    "net_change_in_cash": ("cf_net_change_in_cash",),
    "fx_effect_on_cash": ("cf_fx_effect",),
}

QUARTER_RE = re.compile(r"([1-4])[QT](\d{4})", re.IGNORECASE)
FY_RE = re.compile(r"FY(\d{4})", re.IGNORECASE)
CONSOLIDATED_LABELS: frozenset[str] = frozenset({"consolidado", "consolidated", "total", ""})
RESTATEMENT_REL_TOL = 0.001
RESTATEMENT_ABS_TOL = 0.5


@dataclass(frozen=True)
class Fact:
    source_key: str
    period: str
    basis: str
    value: float
    source_doc: str
    source_ref: str

    @property
    def doc_rank(self) -> tuple[int, int, str]:
        """Filing date proxy from the file name: later filings rank higher."""
        match = QUARTER_RE.search(self.source_doc)
        if match:
            return (int(match.group(2)), int(match.group(1)), self.source_doc)
        match = FY_RE.search(self.source_doc)
        if match:
            return (int(match.group(1)), 5, self.source_doc)
        return (0, 0, self.source_doc)


@dataclass(frozen=True)
class CaptureRow:
    line_item: str
    period: str
    value: float
    source_doc: str
    source_ref: str
    tag: str
    note: str


def ascii_safe(text: str) -> str:
    return text.encode("ascii", "replace").decode("ascii")


def load_facts(inputs_dir: Path) -> list[Fact]:
    files = sorted(inputs_dir.glob("extract_*.json"))
    if not files:
        raise SystemExit(f"[x] no extract_*.json in {ascii_safe(str(inputs_dir))}")
    facts: list[Fact] = []
    for file in files:
        rows = json.loads(file.read_text(encoding="utf-8"))
        if not isinstance(rows, list):
            print(f"[..] skip {file.name}: not a list of facts")
            continue
        kept = 0
        for row in rows:
            if not isinstance(row, dict) or row.get("tag", "observado") not in {"observado", "observed"}:
                continue
            # Segment rows (e.g. region = mexico) would silently replace consolidated figures.
            if str(row.get("region", "consolidado")).lower() not in CONSOLIDATED_LABELS:
                continue
            try:
                value = float(row["value"])
            except (KeyError, TypeError, ValueError):
                continue
            if math.isnan(value):
                continue
            facts.append(Fact(str(row.get("line_item", "")), str(row.get("period", "")), str(row.get("basis", "")),
                              value, str(row.get("source_doc", "")), str(row.get("source_ref", ""))))
            kept += 1
        print(f"[ok] {file.name}: {kept} observed facts")
    return facts


def normalize(item_key: str, value: float) -> float:
    return abs(value) if BY_KEY[item_key].sign is Sign.MAGNITUDE else value


def pick_latest(candidates: list[Fact], restatements: list[dict[str, str]], item_key: str) -> Fact:
    ordered = sorted(candidates, key=lambda f: f.doc_rank)
    chosen = ordered[-1]
    logged: set[tuple[str, str, str]] = {(r["line_item"], r["period"], r["other_value"]) for r in restatements}
    for other in ordered[:-1]:
        diff = abs(abs(other.value) - abs(chosen.value))
        event = (item_key, chosen.period, f"{other.value:.3f}")
        if event in logged:  # same superseded value repeated in several filings: one event
            continue
        if diff > max(RESTATEMENT_ABS_TOL, RESTATEMENT_REL_TOL * abs(chosen.value)):
            logged.add(event)
            restatements.append({
                "line_item": item_key, "period": chosen.period, "basis": chosen.basis,
                "kept_value": f"{chosen.value:.3f}", "kept_doc": chosen.source_doc,
                "other_value": f"{other.value:.3f}", "other_doc": other.source_doc,
                "diff_pct": f"{diff / abs(chosen.value):.4%}" if chosen.value else "n.m.",
            })
    return chosen


def select_rows(facts: list[Fact], years: list[int], scale: float) -> tuple[list[CaptureRow], list[dict[str, str]]]:
    index: dict[tuple[str, str, str], list[Fact]] = defaultdict(list)
    for fact in facts:
        index[(fact.source_key, fact.period, fact.basis)].append(fact)
    rows: list[CaptureRow] = []
    restatements: list[dict[str, str]] = []
    for item in ITEMS:
        source_keys = SOURCE_KEYS.get(item.key, ())
        for year in years:
            row = select_one(item.key, item.statement, source_keys, year, index, restatements, scale)
            if row is not None:
                rows.append(row)
    return rows, restatements


def select_one(item_key: str, statement: Statement, source_keys: tuple[str, ...], year: int,
               index: dict[tuple[str, str, str], list[Fact]], restatements: list[dict[str, str]],
               scale: float) -> CaptureRow | None:
    q4 = f"4Q{year}"
    for source_key in source_keys:
        if statement is Statement.BALANCE:
            candidates = index.get((source_key, q4, "point_in_time"), [])
        else:
            candidates = index.get((source_key, q4, "year_to_date"), [])
        if candidates:
            fact = pick_latest(candidates, restatements, item_key)
            note = "" if source_key == item_key else f"mapped from {source_key}"
            return CaptureRow(item_key, fy_period(year), normalize(item_key, fact.value * scale),
                              fact.source_doc, fact.source_ref, "observed", note)
        if statement is not Statement.BALANCE:
            quarters = [index.get((source_key, f"{q}Q{year}", "quarter"), []) for q in (1, 2, 3, 4)]
            if all(quarters):
                picked = [pick_latest(c, restatements, item_key) for c in quarters]
                total = sum(f.value for f in picked) * scale
                refs = " + ".join(f"{f.source_doc} {f.source_ref}" for f in picked)
                return CaptureRow(item_key, fy_period(year), normalize(item_key, total), picked[-1].source_doc,
                                  refs[:900], "derived", f"sum of 1Q-4Q{year} quarterly values ({source_key})")
    return None


def write_capture(rows: list[CaptureRow], target: Path) -> None:
    with target.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerow(CAPTURE_COLUMNS)
        for r in rows:
            writer.writerow([r.line_item, r.period, f"{r.value:.3f}", r.source_doc, r.source_ref, r.tag, r.note])


def write_restatements(items: list[dict[str, str]], target: Path) -> None:
    columns = ["line_item", "period", "basis", "kept_value", "kept_doc", "other_value", "other_doc", "diff_pct"]
    with target.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=columns)
        writer.writeheader()
        writer.writerows(items)


def validate(capture: Path) -> bool:
    errors: list[str] = []
    seen: set[tuple[str, str]] = set()
    periods: set[str] = set()
    present: set[tuple[str, str]] = set()
    with capture.open(encoding="utf-8", newline="") as handle:
        reader = csv.DictReader(handle)
        if tuple(reader.fieldnames or ()) != CAPTURE_COLUMNS:
            errors.append(f"columns must be {','.join(CAPTURE_COLUMNS)}")
        for line, row in enumerate(reader, start=2):
            key, period = row.get("line_item", ""), row.get("period", "")
            if key not in BY_KEY:
                errors.append(f"line {line}: unknown line_item {key!r}")
                continue
            try:
                period_year(period)
            except ValueError as exc:
                errors.append(f"line {line}: {exc}")
                continue
            try:
                value = float(row.get("value", ""))
            except ValueError:
                errors.append(f"line {line}: value not numeric")
                continue
            if math.isnan(value):
                errors.append(f"line {line}: value is NaN")
            if BY_KEY[key].sign is Sign.MAGNITUDE and value < 0:
                errors.append(f"line {line}: {key} is a MAGNITUDE item and must be >= 0")
            if row.get("tag") not in VALID_TAGS:
                errors.append(f"line {line}: tag must be one of {sorted(VALID_TAGS)}")
            if not row.get("source_doc") or not row.get("source_ref"):
                errors.append(f"line {line}: {key} {period} has no citation")
            if (key, period) in seen:
                errors.append(f"line {line}: duplicate {key} {period}")
            seen.add((key, period))
            periods.add(period)
            present.add((key, period))
    missing = [f"{item.key} {p}" for p in sorted(periods) for item in ITEMS
               if item.required and (item.key, p) not in present]
    for err in errors:
        print(f"[x] {ascii_safe(err)}")
    if missing:
        print(f"[..] missing required items ({len(missing)}), will show #N/A until captured:")
        for entry in missing:
            print(f"     - {entry}")
    print(f"[{'ok' if not errors else 'x'}] validation {'PASS' if not errors else 'FAIL'}: "
          f"{len(seen)} rows, periods {', '.join(sorted(periods))}")
    return not errors


def parse_years(spec: str) -> list[int]:
    if "-" in spec:
        start, end = (int(part) for part in spec.split("-", 1))
        return list(range(start, end + 1))
    return [int(part) for part in spec.split(",")]


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description="Build or validate credit-4c capture.csv.")
    parser.add_argument("inputs", type=Path, help="folder with extract_*.json, or capture.csv with --validate")
    parser.add_argument("out", type=Path, nargs="?", help="output folder")
    parser.add_argument("--years", default="", help="e.g. 2021-2025 or 2023,2024,2025")
    parser.add_argument("--scale", type=float, default=1.0, help="multiply extract values (unit change)")
    parser.add_argument("--validate", action="store_true")
    args = parser.parse_args(argv)
    if args.validate:
        return 0 if validate(args.inputs) else 1
    if args.out is None or not args.years:
        parser.error("out and --years are required unless --validate")
    args.out.mkdir(parents=True, exist_ok=True)
    rows, restatements = select_rows(load_facts(args.inputs), parse_years(args.years), args.scale)
    capture = out_path(args.out, CAPTURE_CSV)
    write_capture(rows, capture)
    write_restatements(restatements, out_path(args.out, RESTATEMENTS_CSV))
    print(f"[ok] wrote {len(rows)} rows -> {ascii_safe(str(capture))}")
    print(f"[ok] {len(restatements)} comparative differences -> {RESTATEMENTS_CSV}")
    keys = {r.line_item for r in rows}
    if "lease_principal_paid" in keys and "lease_interest_paid" in keys:
        print("[..] CHECK before Gate 1: lease_principal_paid comes from a 'lease payments' line. If that line "
              "includes interest, capture principal = line - lease_interest_paid (tag derived).")
    if "lease_liabilities_noncurrent" in keys:
        print("[..] CHECK before Gate 1: tie lease liabilities to the lease note; extract labels can combine "
              "leases with other long-term liabilities.")
    return 0 if validate(capture) else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
