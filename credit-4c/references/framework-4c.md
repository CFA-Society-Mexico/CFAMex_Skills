# 4C framework - definitions, defaults, what to doubt

Base: CFA L1 Fixed Income "Risk, Return, and Credit Analysis" (4 Cs, metric bands),
CFA L2 "Credit Analysis and CDS" (EL = PD x LGD x EAD, ratings lag), S&P Corporate
Methodology "Ratios and Adjustments" (2019-04 criteria: adjusted debt, FFO, FOCF, DCF).

## Contents
1. Metric definitions (what the workbook computes)
2. Policy defaults and their provenance
3. What to doubt - per C (mandatory)
4. Character checklist (keys used in qualitative.json)
5. Coverage rules

---

## 1. Metric definitions

All flows are fiscal-year; balances are fiscal year-end.

| Metric | Formula | Note |
|---|---|---|
| EBITDA (computed) | EBIT + D&A | Used everywhere. Reported EBITDA only feeds a reconciliation check |
| Reported financial debt | short_term_debt + long_term_debt | Bank, bonds, private placements |
| Adjusted gross debt | financial debt + leases x include_leases + pension deficit + receivables factored + supplier finance + guarantees - hybrid equity credit | S&P principle: debt-like obligations count even if not reported as debt |
| Accessible cash | (cash + ST investments - restricted cash) x (1 - haircut) x netting_allowed | S&P: no netting for weak business or sponsor-owned |
| Adjusted net debt | adjusted gross debt - accessible cash | Can be negative (net cash) |
| FFO | EBITDA - cash interest paid - lease interest paid - cash tax paid | |
| CFO after interest | CFO - (interest paid + lease interest) x interest_paid_in_financing | Equals CFO when the issuer books interest in operating |
| CFO before interest | CFO after interest + interest paid + lease interest | Numerator of debt-service coverage |
| FOCF | CFO after interest - capex | |
| FOCF after leases | FOCF - lease principal paid | Matters under IFRS 16: lease principal sits in financing |
| DCF (discretionary) | FOCF after leases - dividends - buybacks | Not discounted cash flow |
| Debt / EBITDA | adjusted gross debt / EBITDA | Core S&P ratio |
| Net debt / EBITDA | adjusted net debt / EBITDA | |
| FFO / debt | FFO / adjusted gross debt | Core S&P ratio; n.m. if debt = 0 |
| EBIT / interest | EBIT / interest expense | interest expense = total financial expense (conservative: includes lease interest, FX) |
| EBITDA / interest | EBITDA / interest expense | |
| Fixed-charge coverage | (EBITDA - capex) / (interest expense + debt due 12m) | debt due 12m = short_term_debt + lease_liabilities_current |
| Debt-service coverage | CFO before interest / (cash interest + lease interest + debt due 12m) | |
| CFO / debt | CFO after interest / adjusted gross debt | |
| Current ratio | current assets / current liabilities | Sector override for retail (negative WC by design) |
| Quick ratio | (current assets - inventories) / current liabilities | |
| Liquidity months | (accessible cash + undrawn committed lines) / ((debt due 12m + cash interest + lease interest) / 12) | Months of debt service covered without new funding |
| Liquidity sources/uses (12m) | (accessible cash + undrawn lines + FFO) / (debt due 12m + capex + dividends) | Latest year as proxy for next 12m |
| Distributions / FOCF after leases | (dividends + buybacks) / FOCF after leases | > 100% = paying shareholders beyond free cash flow |

Ratio cells show `n.m.` when the denominator is <= 0 and `#N/A` when an input
is missing. Status = Approve / Review / Reject by policy band; `n.m.` and missing
data never get a status.

## 2. Policy defaults and their provenance

The user's own policy note has placeholders, so `assets/policy-default.yaml`
ships defaults. Each carries its source so nobody mistakes a default for policy.

| Metric | Approve | Reject | Source |
|---|---|---|---|
| Debt / EBITDA | <= 2.0x | > 3.5x | CFA L1 credit metrics table |
| Net debt / EBITDA | <= 1.5x | > 3.0x | CFA L1 credit metrics table |
| EBIT / interest | >= 5.0x | < 2.5x | CFA L1 credit metrics table |
| EBITDA / interest | >= 7.0x | < 3.5x | CFA L1 credit metrics table |
| Fixed-charge coverage | >= 2.0x | < 1.2x | CFA L1 credit metrics table |
| Debt-service coverage | >= 1.5x | < 1.0x | CFA L1 credit metrics table |
| CFO / debt | >= 30% | < 15% | CFA L1 liquidity table |
| Current ratio | >= 1.0x | < 0.8x | CFA L1 (>1 healthy); reject band is a skill default |
| Quick ratio | >= 0.75x | < 0.5x | CFA L1 (>0.75 healthy); reject band is a skill default |
| FFO / debt | >= 45% | < 20% | Skill default, S&P-style bands from general knowledge - not in the source criteria, verify |
| Liquidity months | >= 24 | < 12 | Skill default |
| Liquidity sources/uses | >= 1.5x | < 1.0x | Skill default |

Note: the CFA L1 text gives Debt/EBITDA "< 3x strong" in prose and "< 2.0x" in its
table. The table is used. Stress: EBITDA -30%, rates +200 bp on the floating share of adjusted debt - taken from the
debt schedule's `rate_type` when a schedule exists (fixed-rate leases don't reprice), else
the policy fallback (100% = conservative).

## 3. What to doubt - per C (mandatory)

Answer every item in `qualitative.json` -> `doubts` with status `resolved` or `open`
and the evidence. An LLM's typical miss is not a wrong number but a question never
asked; this list is the asking.

**Capacity**
- Does the company's "adjusted EBITDA" reconcile to EBIT + D&A? What is added back, and is it recurring?
- Is interest expense complete (lease interest, capitalized interest, FX on debt)?
- Are there debt-like items off the debt line: supplier finance / reverse factoring, receivables sold, pensions, guarantees, earn-outs?
- Is working capital flattering CFO this year (payables stretched, inventory cut)?
- Do lease liabilities tie to the lease note? Is the lease-payment line principal only? Is interest paid booked in operating or financing?
- Undrawn lines: only **committed** facilities count. "Lineas contratadas / no dispuestas" are not committed unless the note says so (comprometidas, with a commitment fee); otherwise capture 0 and say why.
- Is the latest year representative, or is there a one-off (asset sale, tax refund, pandemic year)?
- Sector: is a "weak" ratio actually normal for this business model (retail current ratio < 1)?

**Collateral**
- What ranks ahead of the bonds/loans: secured debt, leases, trade payables, tax and labor claims?
- Are the book values realistic in distress (goodwill = 0, specialized PP&E, inventory obsolescence)?
- Is cash trapped (restricted, abroad, at subsidiaries with minorities)?
- Structural subordination: is the debt at the holding company while assets sit in subsidiaries?

**Covenants**
- Is there any covenant documentation in the folder at all? If not, coverage = none.
- Cure rights: equity cures, broad EBITDA definitions with add-backs, step-ups in limits?
- Covenant-lite or incurrence-only (tested only on new debt)?
- Change-of-control, cross-default, negative pledge, restricted-payment baskets?
- Headroom under the stress case, not just today.

**Character**
- Who controls the company, and do they upstream cash (dividends, royalties, related-party fees) ahead of creditors?
- Restatements, auditor changes, qualified opinions, late filings?
- Track record with creditors: defaults, distressed exchanges, waivers?
- Capital allocation: distributions vs free cash flow; debt-funded buybacks or M&A?
- Disclosure: do the numbers in the press release match the statements? Are segment and debt notes complete?
- Is management's narrative contradicted by the numbers anywhere (margins, cash conversion, guidance misses)?

## 4. Character checklist

Use these `item` keys in `qualitative.json` -> `character` (add others freely):

| key | Question |
|---|---|
| `control_and_governance` | Ownership, controlling shareholder, board independence |
| `related_party` | Related-party transactions, royalties, fees to parent |
| `accounting_quality` | Policies, restatements, adjusted-EBITDA reconciliation |
| `audit` | Auditor, opinion, changes |
| `payment_history` | Defaults, restructurings, waivers |
| `capital_allocation` | Dividends, buybacks, M&A vs free cash flow |
| `disclosure_quality` | Completeness and consistency of disclosure |
| `management_track_record` | Guidance vs delivery, tenure, strategy shifts |
| `creditor_alignment` | Does management act for creditors or only equity |

Flags: `red` (a credit concern), `amber` (watch), `green` (evidence it's fine),
`unknown` (looked, couldn't tell - say where). Need at least 3 `amber`/`red`
concerns overall, or for each `green` an explicit "searched <docs> for <terms>".
Then write a bear case: the 3-5 things that would most plausibly make this credit
go wrong, each tied to a number in the workbook.

## 5. Coverage rules

| C | full | partial | none |
|---|---|---|---|
| Capacity | IS, BS, CF complete for >= 3 years | Gaps in interest, tax, or leases | < 2 years of statements |
| Collateral | Debt by seniority + asset detail | Seniority inferred from notes | No debt or asset detail |
| Covenants | Contract text with definitions | Summary in notes only | Nothing in the folder |
| Character | Annual reports + governance + transcripts | Statements + some notes | Statements only |
