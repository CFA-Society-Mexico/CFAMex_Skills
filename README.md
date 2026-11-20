# Skill

Agent skills for investment analysis, published by CFA Society Mexico. Each folder is a
self-contained skill (`SKILL.md` + references + scripts) that Claude Code and other agents
following the Agent Skills format can load.

| Skill | What it does |
|---|---|
| [credit-4c](#credit-4c) | CFA 4C credit analysis → Excel workbook + PDF report |
| [ips-design](#ips-design) | Individual client IPS from questionnaire answers → Spanish .docx |

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

## ips-design

Drafts an individual client's Investment Policy Statement (IPS) following CFA Institute
practice — return and risk objectives, ability vs willingness to take risk, the
time-horizon / tax / liquidity / legal / unique constraints, strategic allocation and
review triggers — from the client's answers to an intake questionnaire. Deliverable in
Spanish.

**Output**
- A Spanish `.docx` IPS with a draft mark on every page, a signature block and annexes:
  assumptions with sources, open items, and a traceability table linking every client
  fact to the sections that use it.
- The working files behind it: `client.json` (facts + quotes), `calc.json`
  (calculations), `ips.json` (narrative).

**Design principle: the skill drafts; the PM and the client decide.** It is built against
the known failure modes of LLM advisers, chiefly heuristic collapse (letting the stated
risk tolerance drive everything):
- Ability to take risk is scored factor by factor from hard facts (horizon, withdrawal
  rate, liquidity needs, required return, income stability, debt, flexibility).
- Willingness is scored from three independent sources, with contradictions flagged.
- The proposed tolerance is the lower of the two, with an empty PM-decision line.
- Every number comes from a script and reaches the text through placeholders.
- Unknown facts stay `[PENDIENTE]`; tax rates and regulation are never filled from memory.
- The build fails if any client fact is neither used nor justified.

### Install

Copy the `ips-design` folder into your skills directory, e.g. `~/.claude/skills/ips-design/`.

Requirements: Python 3.10+ with `python-docx` (and `pytest` to run the tests).

### Use

- No answers yet: ask the agent for the blank questionnaire
  (`ips-design/assets/questionnaire.md`).
- With answers: e.g. *"Arma la IPS de <cliente>; sus respuestas estan en <archivo>"*. Add
  your capital market expectations (`assets/cma-template.json` format) to get a proposed
  strategic allocation; without them the allocation is left to the PM.
- The skill stops for approval after mapping the answers, and hands over at the end with
  the decisions left open.
- `build_ips_docx.py <out> --final` removes the draft mark, and only runs with zero open
  items.

Default bands in `assets/policy-default.json` are house rules to calibrate, not CFA
figures. Jurisdiction-specific requirements (e.g. the regulator's investor-profile rules)
are left as a compliance checklist.

### Known limitations (v1)

- One risk tolerance for the whole portfolio. There are no goal-based buckets, so a large
  near-term goal (e.g. a home down payment) caps the tolerance of the long-term money too.
- A large uncertain inflow, such as the sale of a business, can't be modeled as a
  scenario; it is kept out of the cash flows and flagged.
- The desired return only varies the end-of-horizon target, not intermediate goals.
- Signature lines are generated for the client and the adviser only; add co-decision
  makers in the text.

### Not for

Institutional investors (pensions, foundations, endowments, insurers), security
selection, or compliance sign-off.

## License

MIT - see [LICENSE](LICENSE).
