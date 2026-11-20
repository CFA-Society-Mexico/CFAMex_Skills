# Skill

Agent skills for investment analysis, published by CFA Society Mexico. Each folder is a
self-contained skill (`SKILL.md` + references + scripts) that Claude Code and other agents
following the Agent Skills format can load.

## credit-4c

Credit analysis of a company from a folder of financial statements, following the CFA
4Cs (Capacity, Collateral, Covenants, Character) with S&P-style debt adjustments, a stress
test and a liquidity / maturity review.

**Output**
- An Excel workbook where every number outside the cited `Sources` table is a live
  formula (blue inputs, black calculations, green links), with integrity checks.
- A PDF report with charts, findings per C, a bear case and the limits of the analysis.

**Design principle: the model explains each C; it does not decide the rating.** Every
status in the workbook is a proposal against policy bands; the analyst-decision cells
are left empty. The pipeline is built against known failure modes of LLM analysts:
arithmetic is done by Excel formulas (never the LLM), every hard-coded figure carries a
document + page citation, each C has a mandatory "what to doubt" list, and only core
metrics can drive a Reject proposal.

### Install

Copy the `credit-4c` folder into your skills directory, e.g. `~/.claude/skills/credit-4c/`.

Requirements: Python 3.10+ with `openpyxl`, `pdfplumber`, `matplotlib`, `reportlab`,
`pyyaml`; Microsoft Excel (Windows, via `pywin32`) or LibreOffice to recalculate the
workbook.

### Use

Ask the agent in plain words, e.g. *"Analiza el credito de <emisora> con el framework 4C;
los estados financieros estan en <carpeta>"*. The skill walks through intake, capture
(with an approval gate), qualitative evidence, workbook build, recalculation and checks,
and the report (with a final hand-over gate).

Default policy bands come from the CFA Level I credit metrics tables plus documented
skill defaults (`credit-4c/assets/policy-default.yaml`). Replace them with your own
policy before relying on any proposal.

### Not for

Banks and insurers (capital and liquidity ratios drive them), equity valuation, or
full three-statement forecasting.
