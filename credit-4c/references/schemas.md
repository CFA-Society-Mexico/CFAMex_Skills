# credit-4c file schemas

Line items, sign rules and file names are defined once in `scripts/registry.py`;
this page explains them. If the two ever disagree, the registry wins.

## capture.csv

One row per (line_item, period). Columns:

| column | meaning |
|---|---|
| `line_item` | key from the registry (below) |
| `period` | fiscal year, `FY2025` |
| `value` | number in the workbook's units (e.g. MXN mm). No thousands separators |
| `source_doc` | file name of the filing |
| `source_ref` | page / table / sheet+row, precise enough to find the number in under a minute |
| `tag` | `observed` (read from a filing), `derived` (computed from observed rows, e.g. sum of four quarters - say how in `note`), `analyst` (judgment input, e.g. an estimate of supplier finance) |
| `note` | optional: reclassification, derivation, caveat |

**Signs.** `SIGNED` items keep the reported sign (profits, CFO, net change in cash).
`MAGNITUDE` items are stored positive: expenses, outflows, liabilities, assets
(e.g. capex `38982`, not `-38982`). The validator rejects negatives on MAGNITUDE items.

**Missing vs zero.** A required item that is missing shows `#N/A` in the workbook,
on purpose. Write `0` only when the filing shows zero (e.g. a debt annex with all
zeros) and cite that page. Optional ADJUSTMENT items missing = 0 in formulas and are
listed as "not captured" in the report limits.

### Line items

Required (R) items must exist for every analyzed period or the gap is reported.

| key | stmt | sign | R |
|---|---|---|---|
| revenue | IS | signed | R |
| operating_profit | IS | signed | R |
| depreciation_amortization | IS | magnitude | R |
| ebitda_reported | IS | signed | |
| interest_expense | IS | magnitude | R |
| interest_income | IS | magnitude | |
| pretax_profit | IS | signed | R |
| income_tax_expense | IS | magnitude | R |
| net_income | IS | signed | R |
| cash_and_equivalents | BS | magnitude | R |
| short_term_investments | BS | magnitude | |
| receivables | BS | magnitude | R |
| inventories | BS | magnitude | R |
| total_current_assets | BS | magnitude | R |
| ppe_net | BS | magnitude | R |
| investment_property | BS | magnitude | |
| rou_assets | BS | magnitude | |
| goodwill | BS | magnitude | |
| intangibles_net | BS | magnitude | |
| total_assets | BS | magnitude | R |
| payables | BS | magnitude | R |
| total_current_liabilities | BS | magnitude | R |
| short_term_debt | BS | magnitude | R |
| long_term_debt | BS | magnitude | R |
| lease_liabilities_current | BS | magnitude | |
| lease_liabilities_noncurrent | BS | magnitude | |
| total_equity | BS | signed | R |
| total_liabilities_and_equity | BS | magnitude | R |
| cfo | CF | signed | R |
| capex | CF | magnitude | R |
| cfi, cff | CF | signed | |
| interest_paid | CF | magnitude | R |
| lease_interest_paid | CF | magnitude | |
| income_tax_paid | CF | magnitude | R |
| lease_principal_paid | CF | magnitude | |
| dividends_paid | CF | magnitude | R |
| share_buybacks | CF | magnitude | |
| cash_at_period_start | CF | magnitude | |
| net_change_in_cash | CF | signed | |
| fx_effect_on_cash | CF | signed | |
| restricted_cash, pension_deficit, receivables_factored, supplier_finance, guarantees, hybrid_equity_credit, undrawn_committed_lines | ADJ | magnitude | |

`lease_principal_paid` = principal only. Many issuers report one "payments of lease
liabilities" line that includes interest: then capture principal = that line - lease
interest, tagged `derived`, and say so in `note`.

`short_term_debt` / `long_term_debt` = bank loans, bonds, private placements and
other interest-bearing financial debt. Leases go in the lease lines, never here.

## qualitative.json

```json
{
  "issuer": {"name": "Wal-Mart de Mexico", "ticker": "WALMEX", "currency": "MXN",
             "units": "MXN mm", "sector": "Food & general merchandise retail",
             "cyclical": false, "fiscal_year_end": "12-31"},
  "as_of": "2026-11-19",
  "source_folder": "C:/.../WALMEX",
  "periods": ["FY2022", "FY2023", "FY2024", "FY2025"],
  "debt_schedule": [
    {"instrument": "Lease liabilities", "seniority": "lease", "currency": "MXN",
     "rate_type": "fixed", "amount": 89175.2,
     "buckets": {"y1": 5186.9, "y2": 0, "y3": 0, "y4": 0, "y5": 0, "after": 83988.3},
     "source_doc": "...", "source_ref": "..."}
  ],
  "pd_annual": null,
  "pd_source": "No rating or market-implied PD in the folder",
  "covenants": [
    {"name": "Max net leverage", "metric": "net_debt_ebitda", "operator": "<=",
     "limit": 3.0, "cure_rights": "broad", "source_doc": "...", "source_ref": "...",
     "note": "EBITDA definition allows pro forma synergies"}
  ],
  "character": [
    {"item": "control_and_governance", "finding": "...", "flag": "amber",
     "source_doc": "...", "source_ref": "..."}
  ],
  "doubts": [
    {"c": "capacity", "doubt": "Does adjusted EBITDA reconcile?", "status": "resolved",
     "evidence": "...", "source_doc": "...", "source_ref": "..."}
  ],
  "bear_case": ["... tied to a workbook number ..."],
  "coverage": {
    "capacity":   {"level": "full",    "reason": "..."},
    "collateral": {"level": "partial", "reason": "..."},
    "covenants":  {"level": "none",    "reason": "No indenture or credit agreement in folder"},
    "character":  {"level": "partial", "reason": "..."}
  }
}
```

- `seniority`: `secured`, `senior_unsecured`, `subordinated`, `lease`.
- `rate_type`: `fixed` or `floating`.
- `buckets`: amounts due in years 1-5 after the latest fiscal year end, and `after`.
  They should sum to `amount`; the workbook checks it. If the filing gives no split by
  year (typical for IFRS lease liabilities at present value), put the current portion
  in `y1` and the rest in `after`; the report then labels it "after year 1 (not split)"
  instead of drawing a fake far-dated wall.
- `covenants[].metric`: a metric key from `policy-default.yaml` (`debt_ebitda`,
  `net_debt_ebitda`, `ebit_interest`, `ebitda_interest`, `ffo_debt`, ...) so the
  workbook links the current value by formula. For any other test use `"other"`
  and add `"current_value"` with its source.
- `cure_rights`: `none`, `limited`, `broad`, `unknown`.
- `character[].flag`: `red`, `amber`, `green`, `unknown`.
- Empty `covenants` list is valid; then `coverage.covenants.level` must be `none` or `partial`.

## findings.json

```json
{
  "headline": "One sentence: what a credit committee should take away.",
  "summary": ["3-5 bullets, each with a number from values.json"],
  "per_c": {
    "capacity":   {"proposal": "Approve", "bullets": ["..."]},
    "collateral": {"proposal": "Review",  "bullets": ["..."]},
    "covenants":  {"proposal": "n.a.",    "bullets": ["..."]},
    "character":  {"proposal": "Review",  "bullets": ["..."]}
  },
  "stress": ["..."],
  "liquidity": ["..."],
  "bear_case": ["..."],
  "context": ["Sector / business-model context that changes how a ratio should be read"],
  "limits": ["What was not verified, data gaps, default (not owned) policy bands"]
}
```

`proposal` uses the workbook's words: `Approve`, `Review`, `Reject`, or `n.a.`
when coverage is none. It is the model's proposal, never the decision.
