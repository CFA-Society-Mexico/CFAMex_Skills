"""Recalculate the credit-4c workbook, scan for errors, export values.

  python recalc_and_check.py <out_dir> [--engine auto|excel|libreoffice]

1. Recalculate: Excel via COM (Windows), else LibreOffice headless. Both save the
   computed values into the file, so openpyxl can read them back.
2. Scan every sheet for formula errors. #N/A = data gap (reported, not FAIL);
   any other error (#REF!, #DIV/0!, #VALUE!, #NAME?, #NUM!, #NULL!) = FAIL.
3. Read the workbook's own Checks verdict.
4. Write values.json (every metric in layout.json with its computed value) and
   checks.json. Exit 1 on FAIL.
"""
from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
from pathlib import Path
from typing import Any

from openpyxl import load_workbook

from registry import CHECKS_JSON, LAYOUT_JSON, VALUES_JSON, out_path

ERROR_STRINGS = ("#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#NUM!", "#NULL!", "#SPILL!", "#CALC!")
NA_STRING = "#N/A"


def recalc_excel(xlsx: Path) -> None:
    import win32com.client  # pywin32; Windows only

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        book = excel.Workbooks.Open(str(xlsx.resolve()))
        excel.CalculateFull()
        book.Save()
        book.Close(SaveChanges=False)
    finally:
        excel.Quit()


def recalc_libreoffice(xlsx: Path) -> None:
    soffice = shutil.which("soffice") or shutil.which("soffice.exe")
    if soffice is None:
        raise RuntimeError("LibreOffice (soffice) not found")
    tmp = xlsx.parent / "_recalc"
    tmp.mkdir(exist_ok=True)
    subprocess.run([soffice, "--headless", "--calc", "--convert-to", "xlsx", "--outdir", str(tmp), str(xlsx)],
                   check=True, capture_output=True, timeout=180)
    shutil.move(str(tmp / xlsx.name), str(xlsx))
    tmp.rmdir()


def recalc(xlsx: Path, engine: str) -> str:
    attempts = {"excel": recalc_excel, "libreoffice": recalc_libreoffice}
    order = ["excel", "libreoffice"] if engine == "auto" else [engine]
    failures: list[str] = []
    for name in order:
        try:
            attempts[name](xlsx)
            return name
        except (ImportError, RuntimeError, OSError, subprocess.SubprocessError) as exc:
            failures.append(f"{name}: {type(exc).__name__}: {exc}")
        except Exception as exc:  # COM raises pywintypes.com_error, not importable without pywin32
            if type(exc).__name__ != "com_error":
                raise
            failures.append(f"{name}: com_error: {exc}")
    raise SystemExit("[x] could not recalculate: " + " | ".join(failures))


def scan_errors(xlsx: Path) -> tuple[list[str], list[str]]:
    wb = load_workbook(xlsx, data_only=True, read_only=True)
    errors, gaps = [], []
    try:
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for cell in row:
                    value = cell.value
                    if isinstance(value, str):
                        if value in ERROR_STRINGS:
                            errors.append(f"{ws.title}!{cell.coordinate} {value}")
                        elif value == NA_STRING:
                            gaps.append(f"{ws.title}!{cell.coordinate}")
    finally:
        wb.close()
    return errors, gaps


def uncomputed(xlsx: Path) -> int:
    """Formula cells with no cached value = recalculation did not happen."""
    formulas = load_workbook(xlsx, data_only=False, read_only=True)
    values = load_workbook(xlsx, data_only=True, read_only=True)
    missing = 0
    try:
        for ws_f in formulas.worksheets:
            ws_v = values[ws_f.title]
            for row_f, row_v in zip(ws_f.iter_rows(), ws_v.iter_rows(), strict=True):
                for cf, cv in zip(row_f, row_v, strict=True):
                    if isinstance(cf.value, str) and cf.value.startswith("=") and cv.value is None:
                        missing += 1
    finally:
        formulas.close()
        values.close()
    return missing


def export_values(xlsx: Path, layout: dict[str, Any]) -> dict[str, Any]:
    wb = load_workbook(xlsx, data_only=True)
    out: dict[str, Any] = {"_meta": layout.get("_meta", {})}
    for key, spec in layout.items():
        if key.startswith("_") or not isinstance(spec, dict) or "sheet" not in spec:
            continue
        ws = wb[spec["sheet"]]
        if "cells" in spec:
            out[key] = {"fmt": spec.get("fmt", ""), "values": {p: ws[c].value for p, c in spec["cells"].items()}}
        elif "cell" in spec:
            out[key] = {"fmt": spec.get("fmt", ""), "value": ws[spec["cell"]].value}
        elif "bucket_cells" in spec:
            out[key] = {"has_schedule": spec["has_schedule"],
                        "buckets": {b: ws[c].value for b, c in spec["bucket_cells"].items()},
                        "cushion": {b: ws[c].value for b, c in spec["cushion_cells"].items()}}
    cov = layout.get("cov.rows", {})
    ws = wb["Covenants"]
    out["covenants"] = [
        {"name": ws[f"A{r}"].value, "metric": ws[f"B{r}"].value, "limit": ws[f"D{r}"].value,
         "current": ws[f"E{r}"].value, "headroom": ws[f"F{r}"].value, "status": ws[f"G{r}"].value,
         "stress_headroom": ws[f"I{r}"].value, "cure_rights": ws[f"J{r}"].value}
        for r in range(cov.get("first", 4), cov.get("first", 4) + cov.get("count", 0))
    ]
    ws = wb["Scorecard"]
    out["scorecard"] = [
        {"dimension": ws[f"A{r}"].value, "coverage": ws[f"B{r}"].value, "approve": ws[f"C{r}"].value,
         "review": ws[f"D{r}"].value, "reject": ws[f"E{r}"].value, "proposal": ws[f"F{r}"].value,
         "evidence": ws[f"G{r}"].value, "analyst": ws[f"H{r}"].value}
        for r in range(5, 11)
    ]
    ws = wb["Character"]
    out["character"] = [
        {"item": ws[f"A{r}"].value, "finding": ws[f"B{r}"].value, "flag": ws[f"C{r}"].value, "source": ws[f"D{r}"].value}
        for r in range(4, ws.max_row + 1)
        if ws[f"A{r}"].value and ws[f"C{r}"].value in {"red", "amber", "green", "unknown"}
    ]
    wb.close()
    return out


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description="Recalculate, check and export the credit-4c workbook.")
    parser.add_argument("out", type=Path)
    parser.add_argument("--engine", choices=("auto", "excel", "libreoffice"), default="auto")
    args = parser.parse_args(argv)
    layout = json.loads(out_path(args.out, LAYOUT_JSON).read_text(encoding="utf-8"))
    xlsx = args.out / layout["_meta"]["xlsx"]
    engine = recalc(xlsx, args.engine)
    print(f"[ok] recalculated with {engine}")

    not_computed = uncomputed(xlsx)
    errors, gaps = scan_errors(xlsx)
    values = export_values(xlsx, layout)
    overall = values.get("chk.overall", {}).get("value")
    out_path(args.out, VALUES_JSON).write_text(json.dumps(values, indent=1, default=str), encoding="utf-8")

    verdict = "PASS"
    if not_computed:
        print(f"[x] {not_computed} formula cells have no computed value")
        verdict = "FAIL"
    if errors:
        print(f"[x] {len(errors)} formula errors:")
        for e in errors[:40]:
            print(f"     {e}")
        verdict = "FAIL"
    if overall != "OK":
        print(f"[x] workbook Checks verdict: {overall}")
        verdict = "FAIL"
    print(f"[..] data gaps (#N/A cells): {len(gaps)}" + (f", e.g. {', '.join(gaps[:8])}" if gaps else ""))
    out_path(args.out, CHECKS_JSON).write_text(json.dumps(
        {"verdict": verdict, "engine": engine, "uncomputed": not_computed, "errors": errors,
         "gaps": gaps, "workbook_checks": overall}, indent=1), encoding="utf-8")
    print(f"[{'ok' if verdict == 'PASS' else 'x'}] {verdict}: values -> {VALUES_JSON}, checks -> {CHECKS_JSON}")
    return 0 if verdict == "PASS" else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
