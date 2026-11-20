"""Single source of truth for credit-4c line items, sign rules and file names.

Every other script imports from here, so the capture schema, the workbook rows and
the validator can never drift apart.
"""
from __future__ import annotations

from dataclasses import dataclass
from enum import Enum, auto
from pathlib import Path


class Statement(Enum):
    INCOME = auto()
    BALANCE = auto()
    CASHFLOW = auto()
    ADJUSTMENT = auto()


class Sign(Enum):
    SIGNED = auto()     # keep the reported sign (profits, net flows)
    MAGNITUDE = auto()  # store as a positive amount (expenses, outflows, liabilities)


@dataclass(frozen=True)
class LineItem:
    key: str
    label: str
    statement: Statement
    sign: Sign
    required: bool


ITEMS: tuple[LineItem, ...] = (
    # Income statement
    LineItem("revenue", "Revenue", Statement.INCOME, Sign.SIGNED, True),
    LineItem("operating_profit", "Operating profit (EBIT)", Statement.INCOME, Sign.SIGNED, True),
    LineItem("depreciation_amortization", "Depreciation & amortization", Statement.INCOME, Sign.MAGNITUDE, True),
    LineItem("ebitda_reported", "EBITDA as reported by the issuer", Statement.INCOME, Sign.SIGNED, False),
    LineItem("interest_expense", "Interest / financial expense", Statement.INCOME, Sign.MAGNITUDE, True),
    LineItem("interest_income", "Interest / financial income", Statement.INCOME, Sign.MAGNITUDE, False),
    LineItem("pretax_profit", "Pre-tax profit", Statement.INCOME, Sign.SIGNED, True),
    LineItem("income_tax_expense", "Income tax expense", Statement.INCOME, Sign.MAGNITUDE, True),
    LineItem("net_income", "Net income", Statement.INCOME, Sign.SIGNED, True),
    # Balance sheet
    LineItem("cash_and_equivalents", "Cash and equivalents", Statement.BALANCE, Sign.MAGNITUDE, True),
    LineItem("short_term_investments", "Short-term investments", Statement.BALANCE, Sign.MAGNITUDE, False),
    LineItem("receivables", "Trade and other receivables", Statement.BALANCE, Sign.MAGNITUDE, True),
    LineItem("inventories", "Inventories", Statement.BALANCE, Sign.MAGNITUDE, True),
    LineItem("total_current_assets", "Total current assets", Statement.BALANCE, Sign.MAGNITUDE, True),
    LineItem("ppe_net", "Property, plant & equipment, net", Statement.BALANCE, Sign.MAGNITUDE, True),
    LineItem("investment_property", "Investment property", Statement.BALANCE, Sign.MAGNITUDE, False),
    LineItem("rou_assets", "Right-of-use assets", Statement.BALANCE, Sign.MAGNITUDE, False),
    LineItem("goodwill", "Goodwill", Statement.BALANCE, Sign.MAGNITUDE, False),
    LineItem("intangibles_net", "Intangibles, net", Statement.BALANCE, Sign.MAGNITUDE, False),
    LineItem("total_assets", "Total assets", Statement.BALANCE, Sign.MAGNITUDE, True),
    LineItem("payables", "Trade and other payables", Statement.BALANCE, Sign.MAGNITUDE, True),
    LineItem("total_current_liabilities", "Total current liabilities", Statement.BALANCE, Sign.MAGNITUDE, True),
    LineItem("short_term_debt", "Short-term debt incl. current portion", Statement.BALANCE, Sign.MAGNITUDE, True),
    LineItem("long_term_debt", "Long-term debt", Statement.BALANCE, Sign.MAGNITUDE, True),
    LineItem("lease_liabilities_current", "Lease liabilities, current", Statement.BALANCE, Sign.MAGNITUDE, False),
    LineItem("lease_liabilities_noncurrent", "Lease liabilities, non-current", Statement.BALANCE, Sign.MAGNITUDE, False),
    LineItem("total_equity", "Total equity", Statement.BALANCE, Sign.SIGNED, True),
    LineItem("total_liabilities_and_equity", "Total liabilities and equity", Statement.BALANCE, Sign.MAGNITUDE, True),
    # Cash flow
    LineItem("cfo", "Cash flow from operations", Statement.CASHFLOW, Sign.SIGNED, True),
    LineItem("capex", "Capital expenditure", Statement.CASHFLOW, Sign.MAGNITUDE, True),
    LineItem("cfi", "Cash flow from investing", Statement.CASHFLOW, Sign.SIGNED, False),
    LineItem("cff", "Cash flow from financing", Statement.CASHFLOW, Sign.SIGNED, False),
    LineItem("interest_paid", "Interest paid (cash)", Statement.CASHFLOW, Sign.MAGNITUDE, True),
    LineItem("lease_interest_paid", "Lease interest paid (cash)", Statement.CASHFLOW, Sign.MAGNITUDE, False),
    LineItem("income_tax_paid", "Income tax paid (cash)", Statement.CASHFLOW, Sign.MAGNITUDE, True),
    LineItem("lease_principal_paid", "Lease principal paid", Statement.CASHFLOW, Sign.MAGNITUDE, False),
    LineItem("dividends_paid", "Dividends paid", Statement.CASHFLOW, Sign.MAGNITUDE, True),
    LineItem("share_buybacks", "Share buybacks", Statement.CASHFLOW, Sign.MAGNITUDE, False),
    LineItem("cash_at_period_start", "Cash at start of period", Statement.CASHFLOW, Sign.MAGNITUDE, False),
    LineItem("net_change_in_cash", "Net change in cash (incl. FX)", Statement.CASHFLOW, Sign.SIGNED, False),
    LineItem("fx_effect_on_cash", "FX effect on cash", Statement.CASHFLOW, Sign.SIGNED, False),
    # Analyst adjustments (optional; missing = 0, flagged as not captured)
    LineItem("restricted_cash", "Restricted / trapped cash", Statement.ADJUSTMENT, Sign.MAGNITUDE, False),
    LineItem("pension_deficit", "Pension / post-retirement deficit (after tax)", Statement.ADJUSTMENT, Sign.MAGNITUDE, False),
    LineItem("receivables_factored", "Receivables sold / factored", Statement.ADJUSTMENT, Sign.MAGNITUDE, False),
    LineItem("supplier_finance", "Supplier finance / reverse factoring", Statement.ADJUSTMENT, Sign.MAGNITUDE, False),
    LineItem("guarantees", "Financial guarantees given", Statement.ADJUSTMENT, Sign.MAGNITUDE, False),
    LineItem("hybrid_equity_credit", "Hybrid equity credit", Statement.ADJUSTMENT, Sign.MAGNITUDE, False),
    LineItem("undrawn_committed_lines", "Undrawn committed credit lines", Statement.ADJUSTMENT, Sign.MAGNITUDE, False),
)

BY_KEY: dict[str, LineItem] = {item.key: item for item in ITEMS}

CAPTURE_COLUMNS: tuple[str, ...] = ("line_item", "period", "value", "source_doc", "source_ref", "tag", "note")
VALID_TAGS: frozenset[str] = frozenset({"observed", "derived", "analyst"})

# File names inside an output folder (one definition, read and written through these).
CAPTURE_CSV = "capture.csv"
RESTATEMENTS_CSV = "restatements.csv"
QUALITATIVE_JSON = "qualitative.json"
POLICY_YAML = "policy.yaml"
LAYOUT_JSON = "layout.json"
VALUES_JSON = "values.json"
FINDINGS_JSON = "findings.json"
CHECKS_JSON = "checks.json"


def out_path(out_dir: Path, name: str) -> Path:
    return out_dir / name


def fy_period(year: int) -> str:
    return f"FY{year}"


def period_year(period: str) -> int:
    if not (period.startswith("FY") and period[2:].isdigit() and len(period) == 6):
        raise ValueError(f"period must look like FY2025, got {period!r}")
    return int(period[2:])
