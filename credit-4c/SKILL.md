---
name: credit-4c
description: Credit analysis of a company from a folder of financial statements, following the CFA 4Cs (Capacity, Collateral, Covenants, Character) with S&P-style debt adjustments, stress test and liquidity/maturity review. Produces a professional Excel workbook where every number outside the cited source table is a live formula, plus a PDF report with charts, findings, bear case and limits. Use this skill whenever the user wants to assess creditworthiness, underwrite or review a borrower or bond issuer, run a "4C" / "4 Cs" analysis, compute leverage/coverage/FFO-to-debt from filings, check covenant headroom, or asks things like "analiza el crédito de X", "¿le prestarías a esta empresa?", "credit memo for this issuer", "run credit-4c on this folder" - even if they don't name the framework. Not for equity valuation (DCF/target price) or for building a full 3-statement forecast model.
---

# credit-4c

Turn a folder of financial statements into a credit analysis an analyst can audit
and defend: an Excel workbook (inputs cited, everything else formulas) and a PDF
report with charts.

**Golden rule: the model explains each C; it does not decide the rating.** Every
status the workbook computes is a *proposal*. The final call per C, and overall,
is an empty analyst cell. This is the user's own policy, and it is also the main
defense against the known failure modes of LLM analysts (below).

## Why the pipeline looks like this

Evidence on LLM analysts (CFA Institute / FrontierFinance / AA-AnalystAgent, 2026)
shows three failures that matter for credit work. Each step below exists to block one:

| Failure | What it looks like in credit | Defense in this skill |
|---|---|---|
| **Coherence** - arithmetic and model flow break | Wrong leverage, ratios that don't tie | The LLM never computes a ratio. Python writes formulas; Excel computes. |
| **Attention** - learns vocabulary, not where to look | Accepts management's "adjusted EBITDA", misses cure rights, treats a retailer's negative working capital as distress | Mandatory "what I doubt" list per C (`references/framework-4c.md`), sector overrides in the policy |
| **Skepticism** - a reviewer that never says "not validated" | Character section that only repeats management's narrative | Character needs >= 3 concerns or an explicit "searched X, Y, Z; found nothing" per item, plus a written bear case |
| **Heuristic collapse** - one input drives the answer | A single Debt/EBITDA number becomes the rating | Scorecard keeps the 4 Cs separate; no auto-rating |
| **Invented numbers** | A figure "remembered" instead of read | Every hard-coded value lives in `Sources` with document + page/sheet reference |

## Outputs

Default output folder: `output/credit-4c/<TICKER>_<YYYY-MM-DD>/` if the working
directory has an `output/` folder, otherwise `./credit-4c/<TICKER>_<YYYY-MM-DD>/`.
**Never write inside the source folder** - it holds the user's filings and other
tools own it.

| File | Written by | Purpose |
|---|---|---|
| `capture.csv` | adapter or you | Every hard-coded figure, one row each, with citation |
| `restatements.csv` | adapter | Comparatives that changed between filings (Character evidence) |
| `qualitative.json` | you | Issuer info, debt schedule, covenants, character evidence, doubts, coverage |
| `<TICKER>_credit-4c_<date>_v01.xlsx` | `build_workbook.py` | The analysis |
| `layout.json` | `build_workbook.py` | Where each metric lives in the workbook |
| `values.json` | `recalc_and_check.py` | Recalculated values, read back from Excel |
| `findings.json` | you | Narrative for the PDF, numbers taken from `values.json` only |
| `<TICKER>_credit-4c_<date>.pdf` | `build_report.py` | Charts + findings report |

Deliverables are in English. Scripts live in `scripts/` next to this file; run them
with the same Python that has `openpyxl`, `pdfplumber`, `matplotlib`, `reportlab`,
`pyyaml` (and `pywin32` on Windows for Excel recalculation).

## Procedure

### 0. Intake
Get from the user (ask only for what you can't infer from the folder):
source folder, issuer name and ticker, currency and units (e.g. MXN mm),
sector and whether it is cyclical, which fiscal years to analyze (default: last
5 with data). Copy `assets/policy-default.yaml` into the output folder as
`policy.yaml`; if the user has their own thresholds, edit that copy - never the asset.

Then think about the business model before any number exists: which default
bands would mislead for this issuer? A grocery retailer runs negative working
capital by design (quick ratio ~0.3x is normal); a utility runs high leverage on
regulated cash flows. Put those in `sector_overrides` in `policy.yaml` with a
`source` that says why, and show them at Gate 1. Metrics carry a `tier`: only
`core` metrics (leverage, FFO/debt, interest coverage, liquidity, recovery) can
make a dimension propose Reject; a `supplementary` metric in Reject only raises
Review - so one ill-fitting ratio cannot flip the whole view.

### 1. Inventory
```
python scripts/inventory.py index <source_folder> <out>/work
```
Classifies files, extracts PDF text per page and dumps xlsx sheets to CSV under
`<out>/work/`, writes `inventory.json`. Indexing a large coverage folder (dozens of
transcripts) takes ~25 minutes: index the filings you need first with
`--only "filings|annual|4Q|FY"`, and add transcripts or management decks later
only for Character evidence. Then search instead of reading whole PDFs:
```
python scripts/inventory.py search <out>/work "covenant|obligaciones de hacer|vencimiento"
```

### 2. Capture (numbers)
Target: `capture.csv` per `references/schemas.md` (line items, sign conventions,
fiscal-year periods `FY2025`).

- **Path A - pre-captured extracts exist** (a `model/inputs/extract_*.json` layout,
  e.g. from the research-analyst plugin):
  ```
  python scripts/capture_from_extracts.py <source_folder>/model/inputs <out> --years 2021-2025
  ```
  It picks fiscal-year values by basis (balance at 4Q, flows year-to-date at 4Q,
  else sum of four quarters tagged `derived`), keeps the latest filing's figure,
  logs restatements, and prints which required items are still missing.
  Do **not** use pre-built wide "canonical" CSVs as the source - they carry no
  citations and may have aggregated point-in-time items incorrectly (seen in
  practice: a year-end cash balance equal to the sum of four quarter-end balances).
  Extracts built for equity models usually lack `short_term_debt` /
  `long_term_debt`: read them from the debt breakdown annex or the debt note
  (e.g. BMV `[800001] Desglose de creditos`) and cite the page - if the annex is
  all zeros, capture `0` with that citation, which is different from missing.
- **Path B - raw filings only**: read the statements from `work/` text and write
  `capture.csv` rows yourself, each with `source_doc` and `source_ref` (page,
  table or sheet/row). Use a deterministic read (parse the text/table), never memory.

Before Gate 1, tie three things to the filing notes - extract keys can carry a
combined label, and these are where it bites:
- **Lease liabilities** to the lease note's total (a line like "leases and other
  long-term liabilities" overstates debt).
- **Lease cash flows**: is the "payments of lease liabilities" line principal only,
  or principal + interest? `lease_principal_paid` must be principal only; if the
  line includes interest, capture principal = total - lease interest (`derived`).
- **Where interest paid sits**: IFRS lets issuers book interest paid (incl. lease
  interest) in financing. If so, set `debt_adjustments.interest_paid_in_financing: 1`
  in `policy.yaml`, so FOCF uses CFO *after* interest.

Fill missing required items from the filings (Path B rules), and add optional
adjustment items you find evidence for (supplier finance, factoring, pensions,
guarantees, undrawn committed lines). Then validate:
```
python scripts/capture_from_extracts.py --validate <out>/capture.csv
```

**Gate 1 - show the user, wait for approval:** periods, units, the mapping of
any line where the issuer's presentation differs from the schema, every
reclassification you made, items left missing (they will show as `#N/A`, not 0),
and the restatements found. Build nothing before an explicit OK.

### 3. Qualitative evidence
Write `qualitative.json` (schema in `references/schemas.md`). Work through
`references/framework-4c.md` section by section:

- **Debt schedule**: every financial debt instrument and lease bucket with
  seniority, rate type, amount, maturity. Source: debt breakdown annex, notes.
- **Covenants**: from indentures / credit agreements / debt notes. If the folder
  has none, say so in `coverage` - do not infer covenants from "typical" deals.
- **Character**: one entry per checklist item, each with finding, proposed flag
  and citation. "Not found" is a finding only if you say where you looked.
- **Doubts**: answer every doubt in the framework's "what to doubt" list for each
  C (resolved / open, with evidence). Add issuer-specific ones.
- **Coverage** per C: `full`, `partial` or `none`, with a one-line reason.

### 4. Build the workbook
```
python scripts/build_workbook.py <out>
```
Reads `capture.csv`, `qualitative.json`, `policy.yaml`. Every non-source cell is
a formula; blue = input, black = calculation, green = link. Never patch numbers
into the xlsx by hand - fix the inputs and rebuild.

### 5. Recalculate and check
```
python scripts/recalc_and_check.py <out>
```
Recalculates through Excel (COM) or LibreOffice, scans for formula errors,
runs the integrity checks, writes `values.json`. Exit code non-zero = FAIL:
fix the cause (input, mapping or script) and rebuild. `#N/A` cells are data gaps,
reported separately - list them in the report limits, do not hide them.

### 6. Findings and report
Read `values.json`, then write `findings.json`: executive summary, per-C proposal
with bullets, bear case, context, limits. **Every number you write must come from
`values.json` or `capture.csv`** - quote it, don't recompute it. Then:
```
python scripts/build_report.py <out>
```

**Gate 2 - hand over:** the model's proposal per C and overall, the open doubts,
the data gaps, and the empty analyst-decision cells in `Scorecard`. The user decides.

## Workbook map

| Sheet | Content |
|---|---|
| Cover | Issuer, period, units, color key, overall check, golden rule |
| Policy | Thresholds, stress parameters, haircuts, toggles (all inputs) |
| Sources | Long table of hard-coded figures with citations (the only hard-codes) |
| Financials | IS/BS/CF by fiscal year, pulled from Sources by formula |
| Adjustments | S&P-style adjusted debt, accessible cash, FFO, FOCF, DCF |
| Capacity | Leverage, coverage, cash-flow and liquidity ratios + status vs policy |
| Stress | Base vs stress (EBITDA shock, rate shock) on latest year |
| Liquidity | 12-month sources/uses, maturity wall from debt schedule |
| Collateral | Distressed asset value, recovery waterfall, LGD, expected loss |
| Covenants | Limit, current value, headroom, cure rights |
| Character | Checklist with evidence, formula flags, analyst flag column |
| Scorecard | 4 Cs side by side, model proposal, analyst decision (empty) |
| Checks | Balance, cash tie-out, CF sum, EBITDA reconciliation, duplicates, errors |

## When things don't fit
- **No financial debt** (only leases or nothing): ratios on financial debt read
  0 and `n.m.` where a denominator is zero; leases still count in adjusted debt.
  Say plainly that Covenants coverage is `none` if there are no debt contracts.
- **Banks, insurers**: this framework does not fit (capital and liquidity ratios
  drive them). Stop and tell the user.
- **Quarterly / LTM**: v1 analyzes fiscal years. If the latest fiscal year is
  stale, say so in limits; do not mix an LTM column into the FY series.
