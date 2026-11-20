"""Build the credit-4c workbook from capture.csv + qualitative.json + policy.yaml.

  python build_workbook.py <out_dir>

Only the Sources sheet (cited figures), Policy (thresholds) and the qualitative
tables (debt schedule, covenants, character evidence) hold typed values. Every
other number is a formula, so an auditor can trace any cell back to a citation.
Writes <TICKER>_credit-4c_<as_of>_v01.xlsx and layout.json (metric -> cell map).
"""
from __future__ import annotations

import argparse
import csv
import json
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import yaml
from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.worksheet import Worksheet

from registry import (
    CAPTURE_CSV,
    ITEMS,
    LAYOUT_JSON,
    POLICY_YAML,
    QUALITATIVE_JSON,
    RESTATEMENTS_CSV,
    Statement,
    out_path,
    period_year,
)

# ---------------------------------------------------------------- styles
BLUE, BLACK, GREEN, GREY, WHITE = "0000FF", "000000", "008000", "808080", "FFFFFF"
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
HEADER_FILL = PatternFill("solid", fgColor="1F3864")
BLOCK_FILL = PatternFill("solid", fgColor="D9E1F2")
FMT_AMT = '#,##0;(#,##0);"-"'
FMT_X = '0.00"x";(0.00"x");"-"'
FMT_PCT = '0.0%;(0.0%);"-"'
FMT_NUM = '0.0;(0.0);"-"'
FMT_BY_POLICY = {"x": FMT_X, "pct": FMT_PCT, "num": FMT_NUM}
FIRST_PERIOD_COL = 3  # periods start in column C on every per-period sheet
STATUS_COLORS = {"Approve": "C6EFCE", "OK": "C6EFCE", "Pass": "C6EFCE", "green": "C6EFCE",
                 "Review": "FFEB9C", "Tight": "FFEB9C", "WARN": "FFEB9C", "amber": "FFEB9C",
                 "Reject": "FFC7CE", "Breach": "FFC7CE", "FAIL": "FFC7CE", "Fail": "FFC7CE", "red": "FFC7CE"}
SENIORITIES = ("secured", "senior_unsecured", "subordinated", "lease")
BUCKETS = ("y1", "y2", "y3", "y4", "y5", "after")


def col(i: int) -> str:
    return get_column_letter(i)


@dataclass
class Ctx:
    """Shared state while building: periods, policy cell addresses, metric layout."""
    periods: list[str]
    policy: dict[str, Any]
    qual: dict[str, Any]
    units: str
    pol: dict[str, str] = field(default_factory=dict)       # policy key -> absolute ref
    fin: dict[str, int] = field(default_factory=dict)       # Financials line key -> row
    adj: dict[str, int] = field(default_factory=dict)       # Adjustments key -> row
    cap: dict[str, int] = field(default_factory=dict)       # Capacity metric -> row
    layout: dict[str, Any] = field(default_factory=dict)

    @property
    def last_col(self) -> str:
        return col(FIRST_PERIOD_COL + len(self.periods) - 1)

    def pcols(self) -> list[str]:
        return [col(FIRST_PERIOD_COL + i) for i in range(len(self.periods))]

    def record_series(self, key: str, sheet: str, row: int, fmt: str) -> None:
        self.layout[key] = {"sheet": sheet, "row": row, "fmt": fmt,
                            "cells": {p: f"{c}{row}" for p, c in zip(self.periods, self.pcols(), strict=True)}}

    def record_cell(self, key: str, sheet: str, cell: str, fmt: str = "") -> None:
        self.layout[key] = {"sheet": sheet, "cell": cell, "fmt": fmt}


# ---------------------------------------------------------------- cell helpers
def put(ws: Worksheet, ref: str, value: Any, *, color: str = BLACK, fmt: str = "", bold: bool = False,
        fill: PatternFill | None = None, italic: bool = False, wrap: bool = False, size: int | None = None) -> None:
    cell = ws[ref]
    cell.value = value
    cell.font = Font(color=color, bold=bold, italic=italic, size=size)
    if fmt:
        cell.number_format = fmt
    if fill is not None:
        cell.fill = fill
    if wrap:
        cell.alignment = Alignment(wrap_text=True, vertical="top")


def put_input(ws: Worksheet, ref: str, value: Any, fmt: str = "") -> None:
    put(ws, ref, value, color=BLUE, fmt=fmt, fill=INPUT_FILL)


def is_link(formula: str) -> bool:
    """A formula that only points at another sheet's cell(s) or pulls by lookup = green."""
    body = formula.lstrip("=")
    return ("!" in body and all(op not in body for op in ("+", "-", "*", "/"))) or body.startswith("IF(COUNTIFS(Sources")


def put_formula(ws: Worksheet, ref: str, formula: str, fmt: str = "", bold: bool = False) -> None:
    put(ws, ref, formula, color=GREEN if is_link(formula) else BLACK, fmt=fmt, bold=bold)


def title(ws: Worksheet, text: str, subtitle: str) -> None:
    put(ws, "A1", text, bold=True, size=14)
    put(ws, "A2", subtitle, italic=True, color=GREY)


def header_row(ws: Worksheet, row: int, labels: list[str], start_col: int = 1) -> None:
    for offset, label in enumerate(labels):
        cell = ws.cell(row=row, column=start_col + offset, value=label)
        cell.font = Font(bold=True, color=WHITE)
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", wrap_text=True, vertical="center")


def block(ws: Worksheet, row: int, text: str, width: int) -> None:
    for c in range(1, width + 1):
        ws.cell(row=row, column=c).fill = BLOCK_FILL
    ws.cell(row=row, column=1, value=text).font = Font(bold=True)


def widths(ws: Worksheet, spec: dict[str, float]) -> None:
    for letter, width in spec.items():
        ws.column_dimensions[letter].width = width


def status_colors(ws: Worksheet, rng: str) -> None:
    for word, color in STATUS_COLORS.items():
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=[f'"{word}"'],
                                                      fill=PatternFill("solid", fgColor=color)))


def ratio(num: str, den: str) -> str:
    return f'=IF({den}<=0,"n.m.",({num})/({den}))'


def status_formula(x: str, direction: str, approve: str, reject: str) -> str:
    if direction == "lower":
        band = f'IF({x}<={approve},"Approve",IF({x}>{reject},"Reject","Review"))'
    else:
        band = f'IF({x}>={approve},"Approve",IF({x}<{reject},"Reject","Review"))'
    return f'=IF(ISNUMBER({x}),{band},IF(ISNA({x}),"No data","n.m."))'


# ---------------------------------------------------------------- loading
def load_capture(path: Path) -> list[dict[str, str]]:
    with path.open(encoding="utf-8", newline="") as handle:
        return list(csv.DictReader(handle))


def load_policy(path: Path) -> dict[str, Any]:
    policy = yaml.safe_load(path.read_text(encoding="utf-8"))
    for key, override in (policy.get("sector_overrides") or {}).items():
        if key not in policy["metrics"]:
            raise SystemExit(f"[x] sector override for unknown metric {key}")
        policy["metrics"][key].update(override)
    return policy


# ---------------------------------------------------------------- sheets
def build_policy(wb: Workbook, ctx: Ctx) -> None:
    ws = wb.create_sheet("Policy")
    title(ws, "Policy", "Thresholds and parameters. All inputs. Defaults carry their source - replace with your own policy.")
    header_row(ws, 4, ["Metric key", "Label", "Direction", "Approve at", "Reject at", "Source", "Tier"])
    row = 5
    for key, m in ctx.policy["metrics"].items():
        fmt = FMT_BY_POLICY[m["fmt"]]
        put(ws, f"A{row}", key, color=GREY)
        put(ws, f"B{row}", m["label"])
        put_input(ws, f"C{row}", m["direction"])
        put_input(ws, f"D{row}", m["approve"], fmt)
        put_input(ws, f"E{row}", m["reject"], fmt)
        put(ws, f"F{row}", m.get("source", ""), italic=True, color=GREY)
        put_input(ws, f"G{row}", m.get("tier", "supplementary"))
        ctx.pol[f"{key}.approve"] = f"Policy!$D${row}"
        ctx.pol[f"{key}.reject"] = f"Policy!$E${row}"
        ctx.pol[f"{key}.tier"] = f"Policy!$G${row}"
        row += 1
    put(ws, f"A{row}", "Tier: only 'core' metrics can make a dimension propose Reject; a 'supplementary' Reject counts as Review.",
        italic=True, color=GREY)
    row += 1
    dv = DataValidation(type="list", formula1='"lower,higher"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add(f"C5:C{row - 1}")
    row += 1
    scalars: list[tuple[str, str, Any, str]] = [
        ("stress.ebitda_shock", "Stress: EBITDA decline", ctx.policy["stress"]["ebitda_shock"], FMT_PCT),
        ("stress.rate_shock_bp", "Stress: rate shock (bp)", ctx.policy["stress"]["rate_shock_bp"], "0"),
        ("stress.floating_share", "Stress: share of adjusted debt that reprices", ctx.policy["stress"]["floating_share"], FMT_PCT),
        ("adj.include_leases", "Debt: include lease liabilities (1/0)", ctx.policy["debt_adjustments"]["include_leases"], "0"),
        ("adj.cash_netting_allowed", "Debt: net accessible cash (1/0)", ctx.policy["debt_adjustments"]["cash_netting_allowed"], "0"),
        ("adj.cash_haircut", "Debt: share of cash not accessible", ctx.policy["debt_adjustments"]["cash_haircut"], FMT_PCT),
        ("adj.interest_in_financing", "Cash flow: interest paid booked in financing (1/0)", ctx.policy["debt_adjustments"].get("interest_paid_in_financing", 0), "0"),
        ("col.admin_claims_share", "Collateral: admin / priority claims share", ctx.policy["collateral"]["admin_claims_share"], FMT_PCT),
        ("col.payables_rank", "Collateral: trade payables rank with unsecured (1/0)", ctx.policy["collateral"]["trade_payables_rank_with_unsecured"], "0"),
        ("cov.tight_headroom", "Covenants: headroom below this = Tight", ctx.policy["covenants"]["tight_headroom"], FMT_PCT),
        ("score.reject_trigger", "Scorecard: # Reject metrics to propose Reject", ctx.policy["scorecard"]["reject_count_trigger"], "0"),
        ("score.review_trigger", "Scorecard: # Review metrics to propose Review", ctx.policy["scorecard"]["review_count_trigger"], "0"),
        ("chk.tolerance", f"Checks: absolute tolerance ({ctx.units})", ctx.policy["checks"]["abs_tolerance"], "0.0"),
    ]
    header_row(ws, row, ["Parameter key", "Parameter", "Value"])
    row += 1
    for key, label, value, fmt in scalars:
        put(ws, f"A{row}", key, color=GREY)
        put(ws, f"B{row}", label)
        put_input(ws, f"C{row}", value, fmt)
        ctx.pol[key] = f"Policy!$C${row}"
        row += 1
    row += 1
    header_row(ws, row, ["Haircut key", "Asset class (distress haircut vs book)", "Haircut"])
    row += 1
    for key, value in ctx.policy["collateral"]["haircuts"].items():
        put(ws, f"A{row}", key, color=GREY)
        put(ws, f"B{row}", key.replace("_", " ").capitalize())
        put_input(ws, f"C{row}", value, FMT_PCT)
        ctx.pol[f"haircut.{key}"] = f"Policy!$C${row}"
        row += 1
    widths(ws, {"A": 24, "B": 46, "C": 12, "D": 12, "E": 12, "F": 60})
    ws.freeze_panes = "A5"


def build_sources(wb: Workbook, ctx: Ctx, rows: list[dict[str, str]]) -> None:
    ws = wb.create_sheet("Sources")
    title(ws, "Sources", f"Every hard-coded figure in this workbook, with its citation. Units: {ctx.units}.")
    header_row(ws, 3, ["line_item", "period", "value", "source_doc", "source_ref", "tag", "note"])
    keep = [r for r in rows if r["period"] in ctx.periods]
    for i, r in enumerate(keep, start=4):
        put(ws, f"A{i}", r["line_item"])
        put(ws, f"B{i}", r["period"])
        put_input(ws, f"C{i}", float(r["value"]), FMT_AMT)
        put(ws, f"D{i}", r["source_doc"])
        put(ws, f"E{i}", r["source_ref"])
        put(ws, f"F{i}", r["tag"])
        put(ws, f"G{i}", r.get("note", ""), italic=True, color=GREY)
    last = max(4, 3 + len(keep))
    ctx.layout["_sources"] = {"first": 4, "last": last}
    widths(ws, {"A": 28, "B": 9, "C": 14, "D": 34, "E": 70, "F": 10, "G": 40})
    ws.freeze_panes = "A4"


def src_range(ctx: Ctx, letter: str) -> str:
    s = ctx.layout["_sources"]
    return f"Sources!${letter}${s['first']}:${letter}${s['last']}"


def build_financials(wb: Workbook, ctx: Ctx) -> None:
    ws = wb.create_sheet("Financials")
    title(ws, "Financials", f"Pulled from Sources by formula. {ctx.units}. #N/A = required item not captured.")
    header_row(ws, 3, ["Line item", "key"] + ctx.periods)
    a, b, c = src_range(ctx, "A"), src_range(ctx, "B"), src_range(ctx, "C")
    row = 4
    blocks = ((Statement.INCOME, "Income statement"), (Statement.BALANCE, "Balance sheet"),
              (Statement.CASHFLOW, "Cash flow"), (Statement.ADJUSTMENT, "Analyst adjustment inputs (missing = 0)"))
    for statement, label in blocks:
        block(ws, row, label, FIRST_PERIOD_COL + len(ctx.periods) - 1)
        row += 1
        for item in (i for i in ITEMS if i.statement is statement):
            put(ws, f"A{row}", item.label + ("" if item.required else " (opt.)"))
            put(ws, f"B{row}", item.key, color=GREY)
            # Required items: a gap must be visible (#N/A). Optional items: absent = 0,
            # and Checks lists how many periods were actually captured.
            missing = "NA()" if item.required else "0"
            for p_col in ctx.pcols():
                f = (f"=IF(COUNTIFS({a},$B{row},{b},{p_col}$3)=0,{missing},"
                     f"SUMIFS({c},{a},$B{row},{b},{p_col}$3))")
                put_formula(ws, f"{p_col}{row}", f, FMT_AMT)
            ctx.fin[item.key] = row
            ctx.record_series(f"fin.{item.key}", "Financials", row, "amt")
            row += 1
        row += 1
    block(ws, row, "Derived", FIRST_PERIOD_COL + len(ctx.periods) - 1)
    row += 1
    derived = (
        ("ebitda", "EBITDA (computed = EBIT + D&A)", "{c}{op} + {c}{da}"),
        ("total_liabilities", "Total liabilities (computed = TL&E - equity)", "{c}{tle} - {c}{eq}"),
        ("financial_debt", "Reported financial debt (ST + LT)", "{c}{std} + {c}{ltd}"),
        ("lease_liabilities", "Lease liabilities (current + non-current)", "{c}{llc} + {c}{lln}"),
    )
    refs = {"op": ctx.fin["operating_profit"], "da": ctx.fin["depreciation_amortization"],
            "tle": ctx.fin["total_liabilities_and_equity"], "eq": ctx.fin["total_equity"],
            "std": ctx.fin["short_term_debt"], "ltd": ctx.fin["long_term_debt"],
            "llc": ctx.fin["lease_liabilities_current"], "lln": ctx.fin["lease_liabilities_noncurrent"]}
    for key, label, template in derived:
        put(ws, f"A{row}", label, bold=True)
        put(ws, f"B{row}", key, color=GREY)
        for p_col in ctx.pcols():
            put_formula(ws, f"{p_col}{row}", "=" + template.format(c=p_col, **refs), FMT_AMT, bold=True)
        ctx.fin[key] = row
        ctx.record_series(f"fin.{key}", "Financials", row, "amt")
        row += 1
    widths(ws, {"A": 46, "B": 26, **{c: 13 for c in ctx.pcols()}})
    ws.freeze_panes = "C4"


def fref(ctx: Ctx, key: str, p_col: str) -> str:
    return f"Financials!{p_col}{ctx.fin[key]}"


def build_adjustments(wb: Workbook, ctx: Ctx) -> None:
    ws = wb.create_sheet("Adjustments")
    title(ws, "Adjustments", f"S&P-style adjusted debt and cash-flow measures. {ctx.units}.")
    header_row(ws, 3, ["Measure", "How it is built"] + ctx.periods)
    P = ctx.pol
    lines: list[tuple[str, str, str, str]] = [
        # key, label, description, formula template ({c} = period column, {r[key]} = row in this sheet)
        ("_h1", "Adjusted debt", "", ""),
        ("financial_debt", "Reported financial debt", "Bank + bonds + private placements", "=Financials!{c}" + str(ctx.fin["financial_debt"])),
        ("leases", "Lease liabilities counted", "Leases x include toggle (Policy)", "=Financials!{c}" + str(ctx.fin["lease_liabilities"]) + "*" + P["adj.include_leases"]),
        ("pension", "Pension deficit", "Underfunded post-retirement obligations", "=" + "Financials!{c}" + str(ctx.fin["pension_deficit"])),
        ("factoring", "Receivables factored", "Sold receivables = secured financing", "=" + "Financials!{c}" + str(ctx.fin["receivables_factored"])),
        ("supplier_fin", "Supplier finance", "Payment terms stretched via banks", "=" + "Financials!{c}" + str(ctx.fin["supplier_finance"])),
        ("guarantees", "Guarantees", "Financial guarantees given", "=" + "Financials!{c}" + str(ctx.fin["guarantees"])),
        ("hybrid", "Less: hybrid equity credit", "Equity content of hybrids", "=-" + "Financials!{c}" + str(ctx.fin["hybrid_equity_credit"])),
        ("adj_debt", "Adjusted gross debt", "Sum of the lines above", "=SUM({c}{r[financial_debt]}:{c}{r[hybrid]})"),
        ("gross_cash", "Cash + ST investments - restricted", "Balance sheet liquidity", "=Financials!{c}" + str(ctx.fin["cash_and_equivalents"]) + "+Financials!{c}" + str(ctx.fin["short_term_investments"]) + "-Financials!{c}" + str(ctx.fin["restricted_cash"])),
        ("acc_cash", "Accessible cash", "x (1 - haircut) x netting toggle", "={c}{r[gross_cash]}*(1-" + P["adj.cash_haircut"] + ")*" + P["adj.cash_netting_allowed"]),
        ("net_debt", "Adjusted net debt", "Adjusted gross debt - accessible cash", "={c}{r[adj_debt]}-{c}{r[acc_cash]}"),
        ("debt_due", "Debt due within 12 months", "ST debt + current lease liabilities", "=Financials!{c}" + str(ctx.fin["short_term_debt"]) + "+Financials!{c}" + str(ctx.fin["lease_liabilities_current"]) + "*" + P["adj.include_leases"]),
        ("_h2", "Earnings and cash flow", "", ""),
        ("ebitda", "EBITDA (computed)", "EBIT + D&A", "=Financials!{c}" + str(ctx.fin["ebitda"])),
        ("ebitda_rep", "EBITDA as reported", "Issuer's own figure (may be adjusted)", "=Financials!{c}" + str(ctx.fin["ebitda_reported"])),
        ("ebitda_gap", "Reported vs computed EBITDA", "Reported / computed - 1", '=IF(AND(ISNUMBER({c}{r[ebitda_rep]}),{c}{r[ebitda_rep]}<>0),IF({c}{r[ebitda]}=0,"n.m.",{c}{r[ebitda_rep]}/{c}{r[ebitda]}-1),"n.a.")'),
        ("interest_exp", "Interest / financial expense", "Income statement", "=Financials!{c}" + str(ctx.fin["interest_expense"])),
        ("int_paid", "Cash interest paid", "Cash flow statement", "=Financials!{c}" + str(ctx.fin["interest_paid"])),
        ("lease_int", "Lease interest paid", "Cash flow statement", "=Financials!{c}" + str(ctx.fin["lease_interest_paid"])),
        ("tax_paid", "Cash tax paid", "Cash flow statement", "=Financials!{c}" + str(ctx.fin["income_tax_paid"])),
        ("ffo", "FFO", "EBITDA - cash interest - lease interest - cash tax", "={c}{r[ebitda]}-{c}{r[int_paid]}-{c}{r[lease_int]}-{c}{r[tax_paid]}"),
        ("cfo", "Cash flow from operations", "Cash flow statement", "=Financials!{c}" + str(ctx.fin["cfo"])),
        ("capex", "Capex", "Cash flow statement", "=Financials!{c}" + str(ctx.fin["capex"])),
        ("cfo_post", "CFO after interest", "CFO - interest paid if the issuer books it in financing (Policy toggle)", "={c}{r[cfo]}-({c}{r[int_paid]}+{c}{r[lease_int]})*" + P["adj.interest_in_financing"]),
        ("cfo_pre", "CFO before interest", "CFO after interest + interest paid + lease interest", "={c}{r[cfo_post]}+{c}{r[int_paid]}+{c}{r[lease_int]}"),
        ("focf", "FOCF", "CFO after interest - capex", "={c}{r[cfo_post]}-{c}{r[capex]}"),
        ("lease_prin", "Lease principal paid", "Financing cash flow (IFRS 16)", "=Financials!{c}" + str(ctx.fin["lease_principal_paid"])),
        ("focf_l", "FOCF after leases", "FOCF - lease principal", "={c}{r[focf]}-{c}{r[lease_prin]}"),
        ("dividends", "Dividends paid", "Cash flow statement", "=Financials!{c}" + str(ctx.fin["dividends_paid"])),
        ("buybacks", "Share buybacks", "Cash flow statement", "=Financials!{c}" + str(ctx.fin["share_buybacks"])),
        ("dcf", "Discretionary cash flow (DCF)", "FOCF after leases - dividends - buybacks", "={c}{r[focf_l]}-{c}{r[dividends]}-{c}{r[buybacks]}"),
        ("distrib_focf", "Distributions / FOCF after leases", ">100% = paying out more than free cash", '=IF({c}{r[focf_l]}<=0,"n.m.",({c}{r[dividends]}+{c}{r[buybacks]})/{c}{r[focf_l]})'),
    ]
    row = 4
    rows: dict[str, int] = {}
    for key, *_ in lines:
        rows[key] = row
        row += 1
    for key, label, desc, template in lines:
        r = rows[key]
        if key.startswith("_h"):
            block(ws, r, label, FIRST_PERIOD_COL + len(ctx.periods) - 1)
            continue
        bold = key in {"adj_debt", "net_debt", "ffo", "focf_l", "dcf"}
        put(ws, f"A{r}", label, bold=bold)
        put(ws, f"B{r}", desc, italic=True, color=GREY)
        fmt = FMT_PCT if key in {"ebitda_gap", "distrib_focf"} else FMT_AMT
        for p_col in ctx.pcols():
            put_formula(ws, f"{p_col}{r}", template.format(c=p_col, r=rows), fmt, bold=bold)
        ctx.adj[key] = r
        ctx.record_series(f"adj.{key}", "Adjustments", r, "pct" if fmt == FMT_PCT else "amt")
    widths(ws, {"A": 36, "B": 46, **{c: 13 for c in ctx.pcols()}})
    ws.freeze_panes = "C4"


def aref(ctx: Ctx, key: str, p_col: str) -> str:
    return f"Adjustments!{p_col}{ctx.adj[key]}"


CAPACITY_METRICS: tuple[tuple[str, str, str], ...] = (
    # metric key, numerator template, denominator template ({A:key} adjustments, {F:key} financials)
    ("debt_ebitda", "{A:adj_debt}", "{A:ebitda}"),
    ("net_debt_ebitda", "{A:net_debt}", "{A:ebitda}"),
    ("ffo_debt", "{A:ffo}", "{A:adj_debt}"),
    ("ebit_interest", "{F:operating_profit}", "{A:interest_exp}"),
    ("ebitda_interest", "{A:ebitda}", "{A:interest_exp}"),
    ("fixed_charge_cov", "({A:ebitda}-{A:capex})", "{A:interest_exp}+{A:debt_due}"),
    ("debt_service_cov", "{A:cfo_pre}", "{A:int_paid}+{A:lease_int}+{A:debt_due}"),
    ("cfo_debt", "{A:cfo_post}", "{A:adj_debt}"),
    ("current_ratio", "{F:total_current_assets}", "{F:total_current_liabilities}"),
    ("quick_ratio", "({F:total_current_assets}-{F:inventories})", "{F:total_current_liabilities}"),
    ("liquidity_months", "({A:acc_cash}+{F:undrawn_committed_lines})", "({A:debt_due}+{A:int_paid}+{A:lease_int})/12"),
)


def expand(template: str, ctx: Ctx, p_col: str) -> str:
    out = template
    for key, row in ctx.adj.items():
        out = out.replace("{A:" + key + "}", f"Adjustments!{p_col}{row}")
    for key, row in ctx.fin.items():
        out = out.replace("{F:" + key + "}", f"Financials!{p_col}{row}")
    if "{" in out:
        raise ValueError(f"unresolved reference in {template}")
    return out


def build_capacity(wb: Workbook, ctx: Ctx) -> None:
    ws = wb.create_sheet("Capacity")
    title(ws, "Capacity", "Ability to pay. Status on the latest year vs Policy bands. n.m. = denominator <= 0.")
    latest = ctx.last_col
    n = len(ctx.periods)
    c_app, c_rej, c_stat, c_src, c_tier = (col(FIRST_PERIOD_COL + n + k) for k in range(5))
    header_row(ws, 3, ["Metric", "Definition"] + ctx.periods + ["Approve at", "Reject at", f"Status {ctx.periods[-1]}", "Band source", "Tier"])
    row = 4
    for key, num, den in CAPACITY_METRICS:
        m = ctx.policy["metrics"][key]
        fmt = FMT_BY_POLICY[m["fmt"]]
        put(ws, f"A{row}", m["label"])
        put(ws, f"B{row}", f"{num} / {den}".replace("{A:", "").replace("{F:", "").replace("}", ""), italic=True, color=GREY)
        for p_col in ctx.pcols():
            put_formula(ws, f"{p_col}{row}", ratio(expand(num, ctx, p_col), expand(den, ctx, p_col)), fmt)
        put_formula(ws, f"{c_app}{row}", "=" + ctx.pol[f"{key}.approve"], fmt)
        put_formula(ws, f"{c_rej}{row}", "=" + ctx.pol[f"{key}.reject"], fmt)
        put_formula(ws, f"{c_stat}{row}", status_formula(f"{latest}{row}", m["direction"], f"{c_app}{row}", f"{c_rej}{row}"), bold=True)
        put(ws, f"{c_src}{row}", m.get("source", ""), italic=True, color=GREY)
        put_formula(ws, f"{c_tier}{row}", "=" + ctx.pol[f"{key}.tier"])
        ctx.cap[key] = row
        ctx.record_series(f"cap.{key}", "Capacity", row, m["fmt"])
        ctx.record_cell(f"cap.{key}.status", "Capacity", f"{c_stat}{row}")
        ctx.record_cell(f"cap.{key}.approve", "Capacity", f"{c_app}{row}", m["fmt"])
        ctx.record_cell(f"cap.{key}.reject", "Capacity", f"{c_rej}{row}", m["fmt"])
        row += 1
    ctx.layout["_capacity_status_range"] = f"Capacity!${c_stat}$4:${c_stat}${row - 1}"
    ctx.layout["_capacity_tier_range"] = f"Capacity!${c_tier}$4:${c_tier}${row - 1}"
    status_colors(ws, f"{c_stat}4:{c_stat}{row - 1}")
    row += 1
    put(ws, f"A{row}", "EBITDA margin")
    for p_col in ctx.pcols():
        put_formula(ws, f"{p_col}{row}", ratio(f"Adjustments!{p_col}{ctx.adj['ebitda']}", f"Financials!{p_col}{ctx.fin['revenue']}"), FMT_PCT)
    ctx.record_series("cap.ebitda_margin", "Capacity", row, "pct")
    widths(ws, {"A": 34, "B": 52, **{c: 11 for c in ctx.pcols()}, c_app: 11, c_rej: 11, c_stat: 14, c_src: 48, c_tier: 13})
    ws.freeze_panes = "C4"


STRESS_METRICS = ("debt_ebitda", "net_debt_ebitda", "ffo_debt", "ebit_interest", "ebitda_interest",
                  "fixed_charge_cov", "debt_service_cov")


def build_stress(wb: Workbook, ctx: Ctx) -> None:
    ws = wb.create_sheet("Stress")
    L = ctx.last_col
    P = ctx.pol
    title(ws, "Stress", f"{ctx.periods[-1]} under Policy stress: EBITDA shock and rate shock on the floating share of adjusted debt.")
    header_row(ws, 3, ["Line", "How", "Base", "Stress", "Status base", "Status stress", "Tier"])
    A = lambda k: f"Adjustments!{L}{ctx.adj[k]}"  # noqa: E731 - short local alias keeps formulas readable
    F = lambda k: f"Financials!{L}{ctx.fin[k]}"  # noqa: E731
    lines = [
        ("factor", "Scenario factor", "0 = base, 1 = stress (structural)", None),
        ("ebitda", "EBITDA", "x (1 - shock x factor)", "={A}*(1-{shock}*{c}4)".format(A=A("ebitda"), shock=P["stress.ebitda_shock"], c="{c}")),
        ("ebit", "EBIT", "EBIT - EBITDA x shock x factor", "={F}-{A}*{shock}*{c}4".format(F=F("operating_profit"), A=A("ebitda"), shock=P["stress.ebitda_shock"], c="{c}")),
        ("extra_int", "Extra interest", "Adj. debt x floating share (debt schedule if present, else Policy) x bp / 10,000 x factor",
         "={A}*IF(ISNUMBER({sch}),{sch},{fl})*{bp}/10000*{c}4".format(A=A("adj_debt"), sch="Liquidity!" + ctx.layout["liq.floating_share"]["cell"], fl=P["stress.floating_share"], bp=P["stress.rate_shock_bp"], c="{c}")),
        ("interest_exp", "Interest expense", "Base + extra", "=" + A("interest_exp") + "+{c}7"),
        ("cash_int", "Cash interest + lease interest", "Base + extra", "=" + A("int_paid") + "+" + A("lease_int") + "+{c}7"),
        ("adj_debt", "Adjusted gross debt", "Unchanged", "=" + A("adj_debt")),
        ("net_debt", "Adjusted net debt", "Unchanged", "=" + A("net_debt")),
        ("ffo", "FFO", "EBITDA - cash interest - cash tax", "={c}5-{c}9-" + A("tax_paid")),
        ("cfo", "CFO before interest", "Base CFO before interest - EBITDA loss", "=" + A("cfo_pre") + "-" + A("ebitda") + "*" + P["stress.ebitda_shock"] + "*{c}4"),
        ("capex", "Capex", "Unchanged", "=" + A("capex")),
        ("debt_due", "Debt due 12m", "Unchanged", "=" + A("debt_due")),
    ]
    rows = {}
    for i, (key, label, how, template) in enumerate(lines, start=4):
        rows[key] = i
        put(ws, f"A{i}", label)
        put(ws, f"B{i}", how, italic=True, color=GREY)
        if template is None:
            put_input(ws, f"C{i}", 0, "0")
            put_input(ws, f"D{i}", 1, "0")
            continue
        for c in ("C", "D"):
            put_formula(ws, f"{c}{i}", template.replace("{c}", c), FMT_AMT)
    r = 4 + len(lines) + 1
    block(ws, r, "Ratios", 6)
    r += 1
    nums = {
        "debt_ebitda": ("{c}10", "{c}5"), "net_debt_ebitda": ("{c}11", "{c}5"), "ffo_debt": ("{c}12", "{c}10"),
        "ebit_interest": ("{c}6", "{c}8"), "ebitda_interest": ("{c}5", "{c}8"),
        "fixed_charge_cov": ("({c}5-{c}14)", "{c}8+{c}15"), "debt_service_cov": ("{c}13", "{c}9+{c}15"),
    }
    first = r
    for key in STRESS_METRICS:
        m = ctx.policy["metrics"][key]
        fmt = FMT_BY_POLICY[m["fmt"]]
        num, den = nums[key]
        put(ws, f"A{r}", m["label"])
        for c in ("C", "D"):
            put_formula(ws, f"{c}{r}", ratio(num.replace("{c}", c), den.replace("{c}", c)), fmt)
        put_formula(ws, f"E{r}", status_formula(f"C{r}", m["direction"], ctx.pol[f"{key}.approve"], ctx.pol[f"{key}.reject"]), bold=True)
        put_formula(ws, f"F{r}", status_formula(f"D{r}", m["direction"], ctx.pol[f"{key}.approve"], ctx.pol[f"{key}.reject"]), bold=True)
        put_formula(ws, f"G{r}", "=" + ctx.pol[f"{key}.tier"])
        ctx.record_cell(f"stress.{key}.base", "Stress", f"C{r}", m["fmt"])
        ctx.record_cell(f"stress.{key}.stress", "Stress", f"D{r}", m["fmt"])
        ctx.record_cell(f"stress.{key}.status", "Stress", f"F{r}")
        r += 1
    status_colors(ws, f"E{first}:F{r - 1}")
    ctx.layout["_stress_status_range"] = f"Stress!$F${first}:$F${r - 1}"
    ctx.layout["_stress_tier_range"] = f"Stress!$G${first}:$G${r - 1}"
    r += 1
    put(ws, f"A{r}", "Stress verdict (Fail = a core metric reaches Reject)", bold=True)
    put_formula(ws, f"D{r}", f'=IF(COUNTIFS(F{first}:F{r - 2},"Reject",G{first}:G{r - 2},"core")>0,"Fail","Pass")', bold=True)
    status_colors(ws, f"D{r}")
    ctx.record_cell("stress.verdict", "Stress", f"D{r}")
    for key, row_i in rows.items():
        if key != "factor":
            ctx.record_cell(f"stress.line.{key}.base", "Stress", f"C{row_i}", "amt")
            ctx.record_cell(f"stress.line.{key}.stress", "Stress", f"D{row_i}", "amt")
    widths(ws, {"A": 34, "B": 46, "C": 14, "D": 14, "E": 13, "F": 13, "G": 13})
    ws.freeze_panes = "C4"


def build_liquidity(wb: Workbook, ctx: Ctx) -> None:
    ws = wb.create_sheet("Liquidity")
    L = ctx.last_col
    title(ws, "Liquidity", f"Next 12 months, latest year ({ctx.periods[-1]}) as proxy; maturity wall from the debt schedule.")
    A = lambda k: f"Adjustments!{L}{ctx.adj[k]}"  # noqa: E731
    F = lambda k: f"Financials!{L}{ctx.fin[k]}"  # noqa: E731
    header_row(ws, 3, ["Sources and uses (12m)", "How", ctx.units])
    items = [
        ("acc_cash", "Accessible cash", "Adjustments", "=" + A("acc_cash")),
        ("undrawn", "Undrawn committed lines", "Analyst input in Sources", "=" + F("undrawn_committed_lines")),
        ("ffo", "FFO", "Latest year", "=" + A("ffo")),
        ("sources", "Total sources", "", "=SUM(C4:C6)"),
        ("debt_due", "Debt due within 12m", "ST debt + current leases", "=" + A("debt_due")),
        ("capex", "Capex", "Latest year", "=" + A("capex")),
        ("dividends", "Dividends", "Latest year", "=" + A("dividends")),
        ("uses", "Total uses", "", "=SUM(C8:C10)"),
    ]
    for i, (key, label, how, f) in enumerate(items, start=4):
        put(ws, f"A{i}", label, bold=key in {"sources", "uses"})
        put(ws, f"B{i}", how, italic=True, color=GREY)
        put_formula(ws, f"C{i}", f, FMT_AMT, bold=key in {"sources", "uses"})
        ctx.record_cell(f"liq.{key}", "Liquidity", f"C{i}", "amt")
    m = ctx.policy["metrics"]["liquidity_su"]
    put(ws, "A13", m["label"], bold=True)
    put_formula(ws, "C13", ratio("C7", "C11"), FMT_X, bold=True)
    put_formula(ws, "D13", status_formula("C13", m["direction"], ctx.pol["liquidity_su.approve"], ctx.pol["liquidity_su.reject"]), bold=True)
    status_colors(ws, "D13")
    ctx.record_cell("liq.liquidity_su", "Liquidity", "C13", "x")
    ctx.record_cell("liq.liquidity_su.status", "Liquidity", "D13")

    sched = ctx.qual.get("debt_schedule") or []
    r = 16
    bucket_labels = ["Year 1", "Year 2", "Year 3", "Year 4", "Year 5", "After 5"]
    header_row(ws, r, ["Instrument", "Seniority", "Rate", "Amount"] + bucket_labels + ["Bucket check", "Source"])
    r += 1
    first = r
    for inst in sched:
        put_input(ws, f"A{r}", inst.get("instrument", ""))
        put_input(ws, f"B{r}", inst.get("seniority", ""))
        put_input(ws, f"C{r}", inst.get("rate_type", ""))
        put_input(ws, f"D{r}", float(inst.get("amount", 0) or 0), FMT_AMT)
        for k, bucket in enumerate(BUCKETS):
            put_input(ws, f"{col(5 + k)}{r}", float((inst.get("buckets") or {}).get(bucket, 0) or 0), FMT_AMT)
        put_formula(ws, f"K{r}", f"=SUM(E{r}:J{r})-D{r}", FMT_AMT)
        put(ws, f"L{r}", f"{inst.get('source_doc', '')} | {inst.get('source_ref', '')}", italic=True, color=GREY)
        r += 1
    if not sched:
        put(ws, f"A{r}", "No debt schedule in qualitative.json - maturity wall not available.", italic=True, color=GREY)
        r += 1
    last = max(first, r - 1)
    put(ws, f"A{r}", "Total maturities", bold=True)
    for k in range(7):
        letter = col(4 + k)
        put_formula(ws, f"{letter}{r}", f"=SUM({letter}{first}:{letter}{last})", FMT_AMT, bold=True)
    put_formula(ws, f"K{r}", f"=SUM(K{first}:K{last})", FMT_AMT, bold=True)
    total_row = r
    r += 1
    put(ws, f"A{r}", "Cumulative maturities")
    for k in range(6):
        letter = col(5 + k)
        put_formula(ws, f"{letter}{r}", f"=SUM($E${total_row}:{letter}{total_row})", FMT_AMT)
    cum_row = r
    r += 1
    put(ws, f"A{r}", "Cushion: cash + undrawn - cumulative maturities")
    for k in range(6):
        letter = col(5 + k)
        put_formula(ws, f"{letter}{r}", f"=$C$4+$C$5-{letter}{cum_row}", FMT_AMT)
    ctx.layout["liq.schedule"] = {"sheet": "Liquidity", "first": first, "last": last, "total_row": total_row,
                                  "cushion_row": r, "has_schedule": bool(sched),
                                  "bucket_cells": {b: f"{col(5 + k)}{total_row}" for k, b in enumerate(BUCKETS)},
                                  "cushion_cells": {b: f"{col(5 + k)}{r}" for k, b in enumerate(BUCKETS)}}
    ctx.record_cell("liq.schedule_total", "Liquidity", f"D{total_row}", "amt")
    r += 2
    put(ws, f"A{r}", "Floating-rate share of debt schedule (used by Stress)")
    put_formula(ws, f"C{r}", f'=IF(D{total_row}>0,SUMIF(C{first}:C{last},"floating",D{first}:D{last})/D{total_row},"n.a.")', FMT_PCT)
    ctx.record_cell("liq.floating_share", "Liquidity", f"C{r}", "pct")
    ctx.record_cell("liq.bucket_check", "Liquidity", f"K{total_row}", "amt")
    ctx.layout["_sched_range"] = {"seniority": f"Liquidity!$B${first}:$B${last}", "amount": f"Liquidity!$D${first}:$D${last}",
                                  "rate": f"Liquidity!$C${first}:$C${last}"}
    widths(ws, {"A": 40, "B": 24, "C": 14, "D": 13, **{col(5 + k): 11 for k in range(6)}, "K": 12, "L": 60})
    ws.freeze_panes = "A4"


def build_collateral(wb: Workbook, ctx: Ctx) -> None:
    ws = wb.create_sheet("Collateral")
    L = ctx.last_col
    P = ctx.pol
    F = lambda k: f"Financials!{L}{ctx.fin[k]}"  # noqa: E731
    title(ws, "Collateral", f"Distressed-value recovery waterfall, {ctx.periods[-1]}. Haircuts are Policy inputs.")
    header_row(ws, 3, ["Asset", "Book value", "Haircut", "Distressed value"])
    assets = [
        ("Cash", F("cash_and_equivalents"), P["haircut.cash"]),
        ("Receivables", F("receivables"), P["haircut.receivables"]),
        ("Inventories", F("inventories"), P["haircut.inventories"]),
        ("PP&E", F("ppe_net"), P["haircut.ppe"]),
        ("Investment property", F("investment_property"), P["haircut.investment_property"]),
        ("Right-of-use assets", F("rou_assets"), P["haircut.rou_assets"]),
        ("Goodwill + intangibles", F("goodwill") + "+" + F("intangibles_net"), P["haircut.goodwill_intangibles"]),
    ]
    r = 4
    for label, book, haircut in assets:
        put(ws, f"A{r}", label)
        put_formula(ws, f"B{r}", "=" + book, FMT_AMT)
        put_formula(ws, f"C{r}", "=" + haircut, FMT_PCT)
        put_formula(ws, f"D{r}", f"=B{r}*(1-C{r})", FMT_AMT)
        r += 1
    put(ws, f"A{r}", "Total", bold=True)
    put_formula(ws, f"B{r}", f"=SUM(B4:B{r - 1})", FMT_AMT, bold=True)
    put_formula(ws, f"D{r}", f"=SUM(D4:D{r - 1})", FMT_AMT, bold=True)
    tot = r
    r += 1
    put(ws, f"A{r}", "Less: admin / priority claims")
    put_formula(ws, f"D{r}", f"=-D{tot}*{P['col.admin_claims_share']}", FMT_AMT)
    r += 1
    put(ws, f"A{r}", "Value available to creditors", bold=True)
    put_formula(ws, f"D{r}", f"=D{tot}+D{r - 1}", FMT_AMT, bold=True)
    avail = r
    ctx.record_cell("col.value_available", "Collateral", f"D{avail}", "amt")
    ctx.record_cell("col.book_total", "Collateral", f"B{tot}", "amt")
    ctx.record_cell("col.distressed_total", "Collateral", f"D{tot}", "amt")

    s = ctx.layout["_sched_range"]
    r += 2
    header_row(ws, r, ["Claims and recovery", "Claims", "Recovery", "LGD"])
    r += 1
    sec, fin_unsec, lease, pay, sub = r, r + 1, r + 2, r + 3, r + 4
    put(ws, f"A{sec}", "Secured debt (debt schedule)")
    put_formula(ws, f"B{sec}", f'=SUMIF({s["seniority"]},"secured",{s["amount"]})', FMT_AMT)
    put(ws, f"A{sub}", "Subordinated debt (debt schedule)")
    put_formula(ws, f"B{sub}", f'=SUMIF({s["seniority"]},"subordinated",{s["amount"]})', FMT_AMT)
    put(ws, f"A{fin_unsec}", "Senior unsecured financial debt (residual)")
    put_formula(ws, f"B{fin_unsec}", f"=MAX(0,{F('short_term_debt')}+{F('long_term_debt')}-B{sec}-B{sub})", FMT_AMT)
    put(ws, f"A{lease}", "Lease liabilities (rank with unsecured)")
    put_formula(ws, f"B{lease}", f"={F('lease_liabilities_current')}+{F('lease_liabilities_noncurrent')}", FMT_AMT)
    put(ws, f"A{pay}", "Trade payables (x rank toggle)")
    put_formula(ws, f"B{pay}", f"={F('payables')}*{P['col.payables_rank']}", FMT_AMT)
    put_formula(ws, f"C{sec}", f'=IF(B{sec}=0,"n.a.",MIN(1,D{avail}/B{sec}))', FMT_PCT)
    rem1 = f"MAX(0,D{avail}-B{sec})"
    pool = f"(B{fin_unsec}+B{lease}+B{pay})"
    put_formula(ws, f"C{fin_unsec}", f'=IF({pool}=0,"n.a.",MIN(1,{rem1}/{pool}))', FMT_PCT)
    put_formula(ws, f"C{lease}", f"=C{fin_unsec}", FMT_PCT)
    put_formula(ws, f"C{pay}", f"=C{fin_unsec}", FMT_PCT)
    put_formula(ws, f"C{sub}", f'=IF(B{sub}=0,"n.a.",MIN(1,MAX(0,{rem1}-{pool})/B{sub}))', FMT_PCT)
    for rr in (sec, fin_unsec, lease, pay, sub):
        put_formula(ws, f"D{rr}", f'=IF(ISNUMBER(C{rr}),1-C{rr},"n.a.")', FMT_PCT)
    r = sub + 2
    m = ctx.policy["metrics"]["recovery_unsecured"]
    put(ws, f"A{r}", m["label"], bold=True)
    put_formula(ws, f"C{r}", f"=C{fin_unsec}", FMT_PCT, bold=True)
    put_formula(ws, f"D{r}", status_formula(f"C{r}", m["direction"], ctx.pol["recovery_unsecured.approve"], ctx.pol["recovery_unsecured.reject"]), bold=True)
    status_colors(ws, f"D{r}")
    ctx.record_cell("col.recovery_unsecured", "Collateral", f"C{r}", "pct")
    ctx.record_cell("col.recovery_unsecured.status", "Collateral", f"D{r}")
    ctx.record_cell("col.recovery_secured", "Collateral", f"C{sec}", "pct")
    ctx.record_cell("col.recovery_sub", "Collateral", f"C{sub}", "pct")
    for k, rr in (("secured", sec), ("senior_unsecured", fin_unsec), ("leases", lease), ("payables", pay), ("subordinated", sub)):
        ctx.record_cell(f"col.claims.{k}", "Collateral", f"B{rr}", "amt")
    r += 1
    put(ws, f"A{r}", "Asset coverage of adjusted debt")
    put_formula(ws, f"C{r}", ratio(f"D{avail}", f"Adjustments!{L}{ctx.adj['adj_debt']}"), FMT_X)
    ctx.record_cell("col.asset_coverage", "Collateral", f"C{r}", "x")
    r += 2
    header_row(ws, r, ["Expected loss (senior unsecured)", "Value", "", "Source"])
    r += 1
    put(ws, f"A{r}", "PD, annual (input)")
    pd = ctx.qual.get("pd_annual")
    put_input(ws, f"B{r}", pd if pd is not None else None, FMT_PCT)
    put(ws, f"D{r}", ctx.qual.get("pd_source", ""), italic=True, color=GREY)
    pd_row = r
    r += 1
    put(ws, f"A{r}", "LGD")
    put_formula(ws, f"B{r}", f"=D{fin_unsec}", FMT_PCT)
    lgd_row = r
    r += 1
    put(ws, f"A{r}", "EAD (senior unsecured financial debt + leases)")
    put_formula(ws, f"B{r}", f"=B{fin_unsec}+B{lease}", FMT_AMT)
    ead_row = r
    r += 1
    put(ws, f"A{r}", "Expected loss = PD x LGD x EAD", bold=True)
    put_formula(ws, f"B{r}", f'=IF(AND(ISNUMBER(B{pd_row}),ISNUMBER(B{lgd_row})),B{pd_row}*B{lgd_row}*B{ead_row},"n.a.")', FMT_AMT, bold=True)
    ctx.record_cell("col.expected_loss", "Collateral", f"B{r}", "amt")
    ctx.record_cell("col.ead", "Collateral", f"B{ead_row}", "amt")
    widths(ws, {"A": 46, "B": 15, "C": 12, "D": 50})


def build_covenants(wb: Workbook, ctx: Ctx) -> None:
    ws = wb.create_sheet("Covenants")
    L = ctx.last_col
    title(ws, "Covenants", "Headroom = distance to limit as % of limit. Current values link to Capacity / Stress by formula.")
    header_row(ws, 3, ["Covenant", "Metric", "Test", "Limit", f"Current {ctx.periods[-1]}", "Headroom", "Status",
                       "Under stress", "Stress headroom", "Cure rights", "Source / note"])
    covs = ctx.qual.get("covenants") or []
    r = 4
    for cv in covs:
        metric = cv.get("metric", "other")
        put_input(ws, f"A{r}", cv.get("name", ""))
        put_input(ws, f"B{r}", metric)
        put_input(ws, f"C{r}", cv.get("operator", "<="))
        put_input(ws, f"D{r}", float(cv.get("limit", 0)), FMT_X)
        if metric in ctx.cap:
            put_formula(ws, f"E{r}", f"=Capacity!{L}{ctx.cap[metric]}", FMT_X)
        else:
            put_input(ws, f"E{r}", cv.get("current_value"), FMT_X)
        head = f'IF(C{r}="<=",(D{r}-{{v}})/D{r},({{v}}-D{r})/D{r})'
        put_formula(ws, f"F{r}", f'=IF(ISNUMBER(E{r}),{head.format(v=f"E{r}")},"n.a.")', FMT_PCT)
        tight = ctx.pol["cov.tight_headroom"]
        put_formula(ws, f"G{r}", f'=IF(ISNUMBER(F{r}),IF(F{r}<0,"Breach",IF(F{r}<{tight},"Tight","OK")),"n.a.")', bold=True)
        if f"stress.{metric}.stress" in ctx.layout:
            put_formula(ws, f"H{r}", "=Stress!" + ctx.layout[f"stress.{metric}.stress"]["cell"], FMT_X)
            put_formula(ws, f"I{r}", f'=IF(ISNUMBER(H{r}),{head.format(v=f"H{r}")},"n.a.")', FMT_PCT)
        else:
            put(ws, f"H{r}", "n.a.", color=GREY)
            put(ws, f"I{r}", "n.a.", color=GREY)
        put_input(ws, f"J{r}", cv.get("cure_rights", "unknown"))
        put(ws, f"K{r}", f"{cv.get('source_doc', '')} | {cv.get('source_ref', '')} {cv.get('note', '')}".strip(),
            italic=True, color=GREY)
        ctx.record_cell(f"cov.{r}", "Covenants", f"F{r}", "pct")
        r += 1
    if not covs:
        put(ws, f"A{r}", "No covenant documentation found in the source folder.", italic=True)
        put(ws, f"G{r}", "n.a.")
        reason = ((ctx.qual.get("coverage") or {}).get("covenants") or {}).get("reason", "")
        put(ws, f"K{r}", reason, italic=True, color=GREY)
        r += 1
    ctx.layout["_cov_status_range"] = f"Covenants!$G$4:$G${r - 1}"
    ctx.layout["cov.rows"] = {"first": 4, "last": r - 1, "count": len(covs)}
    status_colors(ws, f"G4:G{r - 1}")
    widths(ws, {"A": 34, "B": 18, "C": 6, "D": 9, "E": 12, "F": 11, "G": 10, "H": 12, "I": 12, "J": 12, "K": 70})
    ws.freeze_panes = "A4"


def build_character(wb: Workbook, ctx: Ctx, restatements: list[dict[str, str]]) -> None:
    ws = wb.create_sheet("Character")
    L = ctx.last_col
    title(ws, "Character", "Evidence with citations. Flags are the model's proposal; the analyst flag column is yours.")
    header_row(ws, 3, ["Item", "Finding", "Proposed flag", "Source", "Analyst flag"])
    r = 4
    for ch in ctx.qual.get("character") or []:
        put(ws, f"A{r}", ch.get("item", ""))
        put_input(ws, f"B{r}", ch.get("finding", ""))
        ws[f"B{r}"].alignment = Alignment(wrap_text=True, vertical="top")
        put_input(ws, f"C{r}", ch.get("flag", "unknown"))
        put(ws, f"D{r}", f"{ch.get('source_doc', '')} | {ch.get('source_ref', '')}", italic=True, color=GREY, wrap=True)
        put_input(ws, f"E{r}", None)
        r += 1
    block(ws, r, "Formula-driven indicators", 5)
    r += 1
    put(ws, f"A{r}", "capital_allocation (formula)")
    put(ws, f"B{r}", "Distributions / FOCF after leases, latest year (>100% = paying out beyond free cash)")
    put_formula(ws, f"D{r}", f"=Adjustments!{L}{ctx.adj['distrib_focf']}", FMT_PCT)
    put_formula(ws, f"C{r}", f'=IF(ISNUMBER(D{r}),IF(D{r}>1,"amber","green"),"amber")')
    ctx.record_cell("chr.distrib_focf_flag", "Character", f"C{r}")
    r += 1
    first_c, last_c = col(FIRST_PERIOD_COL), L
    put(ws, f"A{r}", "capital_allocation (cumulative, formula)")
    put(ws, f"B{r}", "All analyzed years: (dividends + buybacks) / FOCF after leases")
    put_formula(ws, f"D{r}", ratio(f"SUM(Adjustments!{first_c}{ctx.adj['dividends']}:{last_c}{ctx.adj['dividends']})+SUM(Adjustments!{first_c}{ctx.adj['buybacks']}:{last_c}{ctx.adj['buybacks']})",
                                   f"SUM(Adjustments!{first_c}{ctx.adj['focf_l']}:{last_c}{ctx.adj['focf_l']})"), FMT_PCT)
    put_formula(ws, f"C{r}", f'=IF(ISNUMBER(D{r}),IF(D{r}>1,"amber","green"),"amber")')
    ctx.record_cell("chr.distrib_focf_cum", "Character", f"D{r}", "pct")
    r += 1
    put(ws, f"A{r}", "accounting_quality (formula)")
    put(ws, f"B{r}", "Reported EBITDA vs EBIT + D&A, latest year (> 5% gap = ask what is added back)")
    put_formula(ws, f"D{r}", f"=Adjustments!{L}{ctx.adj['ebitda_gap']}", FMT_PCT)
    put_formula(ws, f"C{r}", f'=IF(ISNUMBER(D{r}),IF(ABS(D{r})>0.05,"amber","green"),"unknown")')
    ctx.record_cell("chr.ebitda_gap_flag", "Character", f"C{r}")
    r += 1
    put(ws, f"A{r}", "accounting_quality (formula)")
    put(ws, f"B{r}", "Comparative figures that changed between filings (table below)")
    rest_first = r + 3
    rest_last = rest_first + max(len(restatements), 1) - 1
    put_formula(ws, f"D{r}", f"=COUNTA(A{rest_first}:A{rest_last})", "0")
    put_formula(ws, f"C{r}", f'=IF(D{r}>0,"amber","green")')
    ctx.record_cell("chr.restatements_count", "Character", f"D{r}", "num")
    flags_last = r
    status_colors(ws, f"C4:C{flags_last}")
    status_colors(ws, f"E4:E{flags_last}")
    ctx.layout["_chr_flag_range"] = f"Character!$C$4:$C${flags_last}"
    r += 2
    header_row(ws, r, ["Line item", "Period", "Kept value (latest filing)", "Kept doc", "Other value", "Other doc"])
    r += 1
    for item in restatements:
        put(ws, f"A{r}", item["line_item"])
        put(ws, f"B{r}", item["period"])
        put_input(ws, f"C{r}", float(item["kept_value"]), FMT_AMT)
        put(ws, f"D{r}", item["kept_doc"], italic=True, color=GREY)
        put_input(ws, f"E{r}", float(item["other_value"]), FMT_AMT)
        put(ws, f"F{r}", item["other_doc"], italic=True, color=GREY)
        r += 1
    widths(ws, {"A": 34, "B": 70, "C": 14, "D": 44, "E": 14, "F": 34})
    ws.freeze_panes = "A4"


def build_scorecard(wb: Workbook, ctx: Ctx) -> None:
    ws = wb.create_sheet("Scorecard")
    ws.sheet_view.showGridLines = False
    title(ws, "Scorecard", "The model proposes; the analyst decides. Proposals never become a rating on their own.")
    header_row(ws, 4, ["Dimension", "Coverage", "# Approve", "# Review", "# Reject", "Model proposal",
                       "Key evidence", "Analyst decision", "Analyst rationale"])
    cov = ctx.qual.get("coverage") or {}
    rej_t, rev_t = ctx.pol["score.reject_trigger"], ctx.pol["score.review_trigger"]
    cr, sr = ctx.layout["_capacity_status_range"], ctx.layout["_stress_status_range"]
    ct, st = ctx.layout["_capacity_tier_range"], ctx.layout["_stress_tier_range"]
    cvr, chr_ = ctx.layout["_cov_status_range"], ctx.layout["_chr_flag_range"]
    L = ctx.last_col
    liq = ctx.layout["liq.liquidity_su.status"]["cell"]
    colr = ctx.layout["col.recovery_unsecured.status"]["cell"]
    dims = [
        # Reject column counts core metrics only; a supplementary Reject is counted as Review.
        ("Capacity", "capacity", f'=COUNTIF({cr},"Approve")',
         f'=COUNTIF({cr},"Review")+COUNTIFS({cr},"Reject",{ct},"supplementary")', f'=COUNTIFS({cr},"Reject",{ct},"core")',
         f'="Debt/EBITDA "&IFERROR(TEXT(Capacity!{L}{ctx.cap["debt_ebitda"]},"0.00")&"x","no data")&"; FFO/debt "&IFERROR(TEXT(Capacity!{L}{ctx.cap["ffo_debt"]},"0%"),"no data")'),
        ("Capacity under stress", "capacity", f'=COUNTIF({sr},"Approve")',
         f'=COUNTIF({sr},"Review")+COUNTIFS({sr},"Reject",{st},"supplementary")', f'=COUNTIFS({sr},"Reject",{st},"core")',
         '="Stress verdict: "&Stress!' + ctx.layout["stress.verdict"]["cell"]),
        ("Liquidity", "capacity", f'=COUNTIF(Liquidity!{liq},"Approve")', f'=COUNTIF(Liquidity!{liq},"Review")', f'=COUNTIF(Liquidity!{liq},"Reject")',
         '="Sources/uses "&IFERROR(TEXT(Liquidity!' + ctx.layout["liq.liquidity_su"]["cell"] + ',"0.00")&"x","n.m.")'),
        ("Collateral", "collateral", f'=COUNTIF(Collateral!{colr},"Approve")', f'=COUNTIF(Collateral!{colr},"Review")', f'=COUNTIF(Collateral!{colr},"Reject")',
         '="Unsecured recovery "&IFERROR(TEXT(Collateral!' + ctx.layout["col.recovery_unsecured"]["cell"] + ',"0%"),"n.a.")'),
        ("Covenants", "covenants", f'=COUNTIF({cvr},"OK")', f'=COUNTIF({cvr},"Tight")', f'=COUNTIF({cvr},"Breach")',
         f'=COUNTA({cvr})-COUNTIF({cvr},"n.a.")&" covenant(s) tested"'),
        ("Character", "character", f'=COUNTIF({chr_},"green")', f'=COUNTIF({chr_},"amber")', f'=COUNTIF({chr_},"red")',
         f'=COUNTIF({chr_},"red")&" red, "&COUNTIF({chr_},"amber")&" amber, "&COUNTIF({chr_},"unknown")&" unknown"'),
    ]
    r = 5
    for label, cov_key, f_app, f_rev, f_rej, f_key in dims:
        put(ws, f"A{r}", label, bold=True)
        c_info = cov.get(cov_key) or {}
        put_input(ws, f"B{r}", c_info.get("level", "unknown"))
        put_formula(ws, f"C{r}", f_app, "0")
        put_formula(ws, f"D{r}", f_rev, "0")
        put_formula(ws, f"E{r}", f_rej, "0")
        put_formula(ws, f"F{r}", f'=IF(C{r}+D{r}+E{r}=0,"n.a.",IF(E{r}>={rej_t},"Reject",IF(D{r}>={rev_t},"Review","Approve")))', bold=True)
        put_formula(ws, f"G{r}", f_key)
        put_input(ws, f"H{r}", None)
        put_input(ws, f"I{r}", None)
        ctx.record_cell(f"score.{label}", "Scorecard", f"F{r}")
        r += 1
    last = r - 1
    r += 1
    put(ws, f"A{r}", "Overall", bold=True, size=12)
    put_formula(ws, f"F{r}", f'=IF(COUNTIF(F5:F{last},"Reject")>0,"Reject",IF(COUNTIF(F5:F{last},"Review")>0,"Review",IF(COUNTIF(F5:F{last},"Approve")>0,"Approve","n.a.")))', bold=True)
    put(ws, f"G{r}", "Worst of the dimensions above", italic=True, color=GREY)
    put_input(ws, f"H{r}", None)
    put_input(ws, f"I{r}", None)
    ctx.record_cell("score.overall", "Scorecard", f"F{r}")
    ctx.record_cell("score.analyst_overall", "Scorecard", f"H{r}")
    status_colors(ws, f"F5:F{r}")
    status_colors(ws, f"H5:H{r}")
    dv = DataValidation(type="list", formula1='"Approve,Review,Reject"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"H5:H{r}")
    r += 2
    put(ws, f"A{r}", "Coverage reasons", bold=True)
    for key in ("capacity", "collateral", "covenants", "character"):
        r += 1
        put(ws, f"A{r}", key)
        put(ws, f"B{r}", (cov.get(key) or {}).get("reason", ""), italic=True, color=GREY)
    widths(ws, {"A": 24, "B": 12, "C": 10, "D": 10, "E": 10, "F": 15, "G": 44, "H": 16, "I": 50})


def build_checks(wb: Workbook, ctx: Ctx) -> None:
    ws = wb.create_sheet("Checks")
    title(ws, "Checks", "Integrity checks. FAIL anywhere -> Cover shows CHECK. #N/A = data gap, counted separately from errors.")
    header_row(ws, 3, ["Check", "Rule"] + ctx.periods + ["Status"])
    stat_col = col(FIRST_PERIOD_COL + len(ctx.periods))
    tol = ctx.pol["chk.tolerance"]
    F = lambda k, c: f"Financials!{c}{ctx.fin[k]}"  # noqa: E731
    checks = [
        ("Balance sheet balances", "Total assets - TL&E", lambda c: f"={F('total_assets', c)}-{F('total_liabilities_and_equity', c)}"),
        ("Cash tie-out", "Start cash + net change - end cash (n.a. if not captured)",
         lambda c: f'=IF(OR({F("cash_at_period_start", c)}=0,{F("net_change_in_cash", c)}=0),"n.a.",{F("cash_at_period_start", c)}+{F("net_change_in_cash", c)}-{F("cash_and_equivalents", c)})'),
        ("Cash flow sums", "CFO + CFI + CFF + FX - net change (n.a. if not captured)",
         lambda c: f'=IF(OR({F("cfi", c)}=0,{F("net_change_in_cash", c)}=0),"n.a.",{F("cfo", c)}+{F("cfi", c)}+{F("cff", c)}+{F("fx_effect_on_cash", c)}-{F("net_change_in_cash", c)})'),
    ]
    r = 4
    for label, rule, make in checks:
        put(ws, f"A{r}", label)
        put(ws, f"B{r}", rule, italic=True, color=GREY)
        for c in ctx.pcols():
            put_formula(ws, f"{c}{r}", make(c), FMT_AMT)
        rng = f"{col(FIRST_PERIOD_COL)}{r}:{ctx.last_col}{r}"
        # COUNTIFS ignores text ("n.a.") and errors, so only real numeric breaks count.
        put_formula(ws, f"{stat_col}{r}", f'=IF(COUNTIFS({rng},">"&{tol})+COUNTIFS({rng},"<"&-{tol})>0,"FAIL","OK")', bold=True)
        r += 1
    put(ws, f"A{r}", "EBITDA reconciliation (warning)")
    put(ws, f"B{r}", "Reported / computed - 1; > 5% = WARN")
    for c in ctx.pcols():
        put_formula(ws, f"{c}{r}", f"=Adjustments!{c}{ctx.adj['ebitda_gap']}", FMT_PCT)
    rng = f"{col(FIRST_PERIOD_COL)}{r}:{ctx.last_col}{r}"
    put_formula(ws, f"{stat_col}{r}", f'=IF(COUNTIFS({rng},">0.05")+COUNTIFS({rng},"<-0.05")>0,"WARN","OK")', bold=True)
    r += 1
    s = ctx.layout["_sources"]
    put(ws, f"A{r}", "Duplicate source rows")
    put(ws, f"B{r}", "Same line_item + period cited twice")
    a_rng, b_rng = f"Sources!$A${s['first']}:$A${s['last']}", f"Sources!$B${s['first']}:$B${s['last']}"
    put_formula(ws, f"C{r}", f"=SUMPRODUCT((COUNTIFS({a_rng},{a_rng},{b_rng},{b_rng})>1)*1)", "0")
    put_formula(ws, f"{stat_col}{r}", f'=IF(C{r}>0,"FAIL","OK")', bold=True)
    r += 1
    sched = ctx.layout["liq.schedule"]
    put(ws, f"A{r}", "Debt schedule buckets sum to amounts")
    put(ws, f"B{r}", "Sum of bucket differences")
    put_formula(ws, f"C{r}", "=Liquidity!" + ctx.layout["liq.bucket_check"]["cell"], FMT_AMT)
    put_formula(ws, f"{stat_col}{r}", f'=IF(ABS(C{r})>{tol},"FAIL","OK")', bold=True)
    r += 1
    put(ws, f"A{r}", "Debt schedule vs balance sheet (warning)")
    put(ws, f"B{r}", "Schedule total - (financial debt + leases), latest year")
    L = ctx.last_col
    put_formula(ws, f"C{r}", f"=Liquidity!{ctx.layout['liq.schedule_total']['cell']}-(Financials!{L}{ctx.fin['financial_debt']}+Financials!{L}{ctx.fin['lease_liabilities']})", FMT_AMT)
    if sched["has_schedule"]:
        put_formula(ws, f"{stat_col}{r}", f'=IF(ABS(C{r})>{tol},"WARN","OK")', bold=True)
    else:
        put(ws, f"{stat_col}{r}", "n.a.")
    r += 2
    block(ws, r, "Formula errors and data gaps", FIRST_PERIOD_COL + len(ctx.periods))
    r += 1
    err_rows = []
    for sheet, rng in (("Financials", f"C4:{L}200"), ("Adjustments", f"C4:{L}60"), ("Capacity", f"C4:{col(FIRST_PERIOD_COL + len(ctx.periods) + 3)}40"),
                       ("Stress", "C4:F40"), ("Liquidity", "C4:K80"), ("Collateral", "B4:D60"), ("Covenants", "D4:I80"),
                       ("Character", "C4:E200"), ("Scorecard", "C5:G20")):
        put(ws, f"A{r}", f"{sheet}: formula errors (excl. #N/A)")
        put_formula(ws, f"C{r}", f"=SUMPRODUCT(ISERROR({sheet}!{rng})*1)-SUMPRODUCT(ISNA({sheet}!{rng})*1)", "0")
        put(ws, f"D{r}", "data gaps (#N/A)", italic=True, color=GREY)
        put_formula(ws, f"E{r}", f"=SUMPRODUCT(ISNA({sheet}!{rng})*1)", "0")
        err_rows.append(r)
        r += 1
    put(ws, f"A{r}", "Total formula errors", bold=True)
    put_formula(ws, f"C{r}", "=" + "+".join(f"C{x}" for x in err_rows), "0", bold=True)
    put_formula(ws, f"E{r}", "=" + "+".join(f"E{x}" for x in err_rows), "0", bold=True)
    put_formula(ws, f"{stat_col}{r}", f'=IF(C{r}>0,"FAIL","OK")', bold=True)
    ctx.record_cell("chk.errors", "Checks", f"C{r}", "num")
    ctx.record_cell("chk.gaps", "Checks", f"E{r}", "num")
    r += 2
    put(ws, f"A{r}", "Overall", bold=True, size=12)
    put_formula(ws, f"{stat_col}{r}", f'=IF(COUNTIF({stat_col}4:{stat_col}{r - 2},"FAIL")>0,"CHECK","OK")', bold=True)
    ctx.record_cell("chk.overall", "Checks", f"{stat_col}{r}")
    status_colors(ws, f"{stat_col}4:{stat_col}{r}")
    ws.conditional_formatting.add(f"{stat_col}{r}", CellIsRule(operator="equal", formula=['"CHECK"'], fill=PatternFill("solid", fgColor="FFC7CE")))
    r += 2
    block(ws, r, "Optional items captured (periods with a cited value)", FIRST_PERIOD_COL + len(ctx.periods))
    r += 1
    for item in (i for i in ITEMS if not i.required):
        put(ws, f"A{r}", item.label)
        put(ws, f"B{r}", item.key, color=GREY)
        put_formula(ws, f"C{r}", f'=COUNTIF({a_rng},"{item.key}")', "0")
        r += 1
    widths(ws, {"A": 42, "B": 52, **{c: 12 for c in ctx.pcols()}, stat_col: 10})
    ws.freeze_panes = "C4"


def build_cover(wb: Workbook, ctx: Ctx, out_dir: Path) -> None:
    ws = wb["Cover"]
    ws.sheet_view.showGridLines = False
    iss = ctx.qual.get("issuer") or {}
    put(ws, "B2", f"{iss.get('name', '')} ({iss.get('ticker', '')}) - Credit analysis, 4C framework", bold=True, size=16)
    put(ws, "B3", "Model proposes, analyst decides: no status in this workbook is a rating.", italic=True, color=GREY)
    info = [
        ("Sector", iss.get("sector", "")), ("Cyclical", "yes" if iss.get("cyclical") else "no"),
        ("Units", ctx.units), ("Periods", f"{ctx.periods[0]} - {ctx.periods[-1]}"),
        ("As of", ctx.qual.get("as_of", "")), ("Source folder", ctx.qual.get("source_folder", str(out_dir))),
        ("Prepared with", "credit-4c skill (formulas built by script; figures cited in Sources)"),
    ]
    r = 5
    for label, value in info:
        put(ws, f"B{r}", label, bold=True)
        put(ws, f"C{r}", value)
        r += 1
    r += 1
    put(ws, f"B{r}", "Integrity check", bold=True)
    put_formula(ws, f"C{r}", "=Checks!" + ctx.layout["chk.overall"]["cell"], bold=True)
    ws.conditional_formatting.add(f"C{r}", CellIsRule(operator="equal", formula=['"OK"'], fill=PatternFill("solid", fgColor="C6EFCE")))
    ws.conditional_formatting.add(f"C{r}", CellIsRule(operator="equal", formula=['"CHECK"'], fill=PatternFill("solid", fgColor="FFC7CE")))
    r += 1
    put(ws, f"B{r}", "Model proposal (overall)", bold=True)
    put_formula(ws, f"C{r}", "=Scorecard!" + ctx.layout["score.overall"]["cell"], bold=True)
    status_colors(ws, f"C{r}")
    r += 1
    put(ws, f"B{r}", "Analyst decision (overall)", bold=True)
    put_formula(ws, f"C{r}", f'=IF(Scorecard!{ctx.layout["score.analyst_overall"]["cell"]}="","(pending)",Scorecard!{ctx.layout["score.analyst_overall"]["cell"]})', bold=True)
    r += 2
    put(ws, f"B{r}", "Color key", bold=True)
    for label, color, fill in (("Input (hard-coded, cited or policy)", BLUE, INPUT_FILL), ("Calculation", BLACK, None),
                               ("Link / lookup from another sheet", GREEN, None)):
        r += 1
        put(ws, f"B{r}", "1,234", color=color, fill=fill)
        put(ws, f"C{r}", label)
    r += 2
    put(ws, f"B{r}", "Sheet guide", bold=True)
    guide = [("Policy", "Thresholds and parameters"), ("Sources", "Every hard-coded figure with citation"),
             ("Financials", "Statements pulled from Sources"), ("Adjustments", "Adjusted debt, FFO, FOCF, DCF"),
             ("Capacity", "Ratios and status vs policy"), ("Stress", "Base vs stress"), ("Liquidity", "12m sources/uses, maturity wall"),
             ("Collateral", "Recovery waterfall, LGD, EL"), ("Covenants", "Headroom"), ("Character", "Evidence and flags"),
             ("Scorecard", "4 Cs side by side, analyst decision"), ("Checks", "Integrity checks")]
    for sheet, text in guide:
        r += 1
        cell = ws[f"B{r}"]
        cell.value = sheet
        cell.hyperlink = f"#{sheet}!A1"
        cell.font = Font(color="0563C1", underline="single")
        put(ws, f"C{r}", text)
    widths(ws, {"A": 3, "B": 34, "C": 80})


# ---------------------------------------------------------------- main
def build(out_dir: Path) -> Path:
    capture = load_capture(out_path(out_dir, CAPTURE_CSV))
    qual = json.loads(out_path(out_dir, QUALITATIVE_JSON).read_text(encoding="utf-8"))
    policy = load_policy(out_path(out_dir, POLICY_YAML))
    rest_path = out_path(out_dir, RESTATEMENTS_CSV)
    restatements = load_capture(rest_path) if rest_path.exists() else []
    periods = qual.get("periods") or sorted({r["period"] for r in capture}, key=period_year)
    periods = sorted(periods, key=period_year)
    restatements = [x for x in restatements if f"FY{x['period'][-4:]}" in periods]
    iss = qual.get("issuer") or {}
    ctx = Ctx(periods=periods, policy=policy, qual=qual, units=iss.get("units", ""))

    wb = Workbook()
    wb.active.title = "Cover"
    build_policy(wb, ctx)
    build_sources(wb, ctx, capture)
    build_financials(wb, ctx)
    build_adjustments(wb, ctx)
    build_capacity(wb, ctx)
    build_liquidity(wb, ctx)  # before Stress: Stress reads the schedule's floating share
    build_stress(wb, ctx)
    build_collateral(wb, ctx)
    build_covenants(wb, ctx)
    build_character(wb, ctx, restatements)
    build_scorecard(wb, ctx)
    build_checks(wb, ctx)
    build_cover(wb, ctx, out_dir)
    order = ["Cover", "Scorecard", "Capacity", "Stress", "Liquidity", "Collateral", "Covenants", "Character",
             "Adjustments", "Financials", "Sources", "Policy", "Checks"]
    wb._sheets = [wb[name] for name in order]  # noqa: SLF001 - openpyxl has no public reorder API
    wb.calculation.fullCalcOnLoad = True

    ticker = iss.get("ticker", "ISSUER")
    as_of = qual.get("as_of", "")
    xlsx = out_dir / f"{ticker}_credit-4c_{as_of}_v01.xlsx"
    wb.save(xlsx)
    ctx.layout["_meta"] = {"xlsx": xlsx.name, "periods": periods, "units": ctx.units, "ticker": ticker,
                           "policy_metrics": policy["metrics"]}
    out_path(out_dir, LAYOUT_JSON).write_text(json.dumps(ctx.layout, indent=1), encoding="utf-8")
    return xlsx


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description="Build the credit-4c workbook (formulas, not values).")
    parser.add_argument("out", type=Path, help="output folder with capture.csv, qualitative.json, policy.yaml")
    args = parser.parse_args(argv)
    xlsx = build(args.out)
    print(f"[ok] workbook -> {str(xlsx).encode('ascii', 'replace').decode('ascii')}")
    print("[..] values are not computed yet: run recalc_and_check.py next")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
