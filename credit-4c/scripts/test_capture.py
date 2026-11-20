"""pytest for capture_from_extracts: the logic that decides which number enters the workbook."""
from __future__ import annotations

import csv
import json
from pathlib import Path

import pytest

from capture_from_extracts import load_facts, select_rows, validate
from registry import CAPTURE_COLUMNS


def fact(key: str, period: str, value: float, basis: str, doc: str, region: str = "consolidado") -> dict[str, str]:
    return {"line_item": key, "period": period, "value": str(value), "basis": basis, "region": region,
            "statement": "x", "source_doc": doc, "source_ref": "p.1", "tag": "observado"}


@pytest.fixture()
def inputs(tmp_path: Path) -> Path:
    rows = [
        fact("total_assets", "4Q2025", 100.0, "point_in_time", "R_4Q2025.pdf"),
        fact("total_assets", "4Q2025", 101.0, "point_in_time", "R_1Q2026.pdf"),        # later filing restates
        fact("total_assets", "4Q2025", 999.0, "point_in_time", "R_2Q2026.pdf", "mexico"),  # segment: ignored
        fact("capex", "4Q2025", -40.0, "year_to_date", "R_4Q2025.pdf"),
        *[fact("cf_interest_paid", f"{q}Q2025", -q, "quarter", f"X_{q}Q2025.xlsx") for q in (1, 2, 3, 4)],
        fact("cash_and_equivalents", "1Q2025", 7.0, "point_in_time", "R_1Q2025.pdf"),  # not year end: ignored
    ]
    (tmp_path / "extract_a.json").write_text(json.dumps(rows), encoding="utf-8")
    return tmp_path


def test_selection_rules(inputs: Path) -> None:
    rows, restatements = select_rows(load_facts(inputs), [2025], 1.0)
    by_key = {r.line_item: r for r in rows}
    assert by_key["total_assets"].value == 101.0          # latest filing wins
    assert by_key["total_assets"].source_doc == "R_1Q2026.pdf"
    assert by_key["capex"].value == 40.0                   # magnitude item stored positive
    assert by_key["interest_paid"].value == 10.0           # 1+2+3+4 quarters
    assert by_key["interest_paid"].tag == "derived"
    assert "cash_and_equivalents" not in by_key            # no year-end balance -> missing, not 7
    assert len(restatements) == 1 and restatements[0]["other_value"] == "100.000"


def write_capture(path: Path, rows: list[list[str]]) -> Path:
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerow(CAPTURE_COLUMNS)
        writer.writerows(rows)
    return path


def test_validate_rejects_bad_rows(tmp_path: Path) -> None:
    bad = write_capture(tmp_path / "c.csv", [
        ["capex", "FY2025", "-5", "d", "r", "observed", ""],          # negative magnitude
        ["revenue", "FY2025", "10", "", "", "observed", ""],          # no citation
        ["revenue", "FY2025", "10", "d", "r", "observed", ""],        # duplicate
        ["nonsense", "FY2025", "1", "d", "r", "observed", ""],        # unknown key
    ])
    assert validate(bad) is False


def test_validate_passes_clean_rows(tmp_path: Path) -> None:
    good = write_capture(tmp_path / "c.csv", [["revenue", "FY2025", "10", "d.pdf", "p.3", "observed", ""]])
    assert validate(good) is True
