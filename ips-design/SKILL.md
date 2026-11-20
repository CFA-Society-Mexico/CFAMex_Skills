---
name: ips-design
description: Draft a client's Investment Policy Statement (IPS / Declaración de Política de Inversión) following CFA Institute best practice (objectives and constraints, RRTTLLU, ability vs willingness, multistage horizon, review triggers), from the client's answers to an intake questionnaire, and deliver it as a Spanish .docx. Use this skill whenever the user wants to write, draft, build or update an IPS or a client investment policy, turn questionnaire answers / a client interview / meeting notes into a policy, profile an individual investor's return and risk objectives and constraints, or says things like "arma la IPS de este cliente", "redacta la política de inversión", "tengo las respuestas del cuestionario de perfil", "IPS for my new client" — even if they don't say "IPS". Also use it to hand out the blank intake questionnaire. Not for institutional investors (pensions, endowments, insurers), for picking securities, or for compliance sign-off.
---

# ips-design

Turn a client's questionnaire answers into an IPS draft the PM can defend: every
number calculated by a script, every statement traced to a client answer, every gap
marked instead of guessed.

**Golden rule: the skill drafts; the PM and the client decide.** The risk tolerance in
the document is a *proposal* with an empty "Decisión del PM" line; the strategic
allocation comes from the PM's capital market expectations or stays a placeholder; the
document carries a draft mark until the PM removes it.

## Why the pipeline looks like this

| Failure | What it looks like in an IPS | Defense |
|---|---|---|
| **Heuristic collapse** — one input drives the answer | Risk profile = whatever the client ticked on the questionnaire; income, horizon and liquidity ignored | Ability scored factor by factor from hard facts; willingness from three sources; tolerance = min of the two; section 5 shows every factor (`references/risk-tolerance.md`) |
| **Invented facts** | A tax rate, life expectancy or bequest "assumed" silently | Unknown = `null` → `[PENDIENTE]`; assumptions only from `policy.json` with a source, listed in Annex A |
| **Broken arithmetic** | Required return that doesn't tie to the goals | `ips_calc.py` computes; the text quotes `{{calc...}}` placeholders; the builder fails on unresolved ones |
| **Generic boilerplate** | "El cliente tiene horizonte de largo plazo" | Traceability: every client fact must feed a section or be justified in `no_usados`; the build fails otherwise |
| **Looks final when it isn't** | A polished docx that never saw compliance | Draft header on every page; `--final` refuses while any `[PENDIENTE]` remains |

## Outputs

Folder: `output/ips/<slug>_<fecha_ips>/` if the working directory has an `output/`
folder, otherwise `./ips/<slug>_<fecha_ips>/`. Never write next to the client's source
files. Client data stays in that folder — never copy it into notes, wikis or logs.

| File | Written by | Purpose |
|---|---|---|
| `client.json` | you | Every fact from the answers, `null` when unknown, with quotes in `fuentes` |
| `policy.json` | copy of `assets/policy-default.json` | Bands, caps and assumptions (edit the copy, never the asset) |
| `cma.json`, `saa.json` | PM / you | Optional: capital market expectations and proposed allocation |
| `calc.json` | `ips_calc.py` | Required return, liquidity, ability, willingness, tolerance, SAA check |
| `ips.json` | you | The narrative, with placeholders and auto-table blocks |
| `IPS_<slug>_<fecha>.docx` | `build_ips_docx.py` | The deliverable (Spanish) |

Scripts live in `scripts/` next to this file and need `python-docx`. Schemas for every file:
`references/schemas.md`.

## Procedure

### 0. No answers yet?
If the user has a new client but no answers, give them `assets/questionnaire.md` (offer
it as a .docx if they want to send it). The questionnaire is what makes the ability
factors computable; free-form notes work too, but expect more `[PENDIENTE]`.

### 1. Intake → `client.json`
Read the answers (filled questionnaire, email, interview transcript, notes) and map them
to `client.json`. Rules:
- Quote the client for each fact in `fuentes` (question id + short quote).
- Unknown, "no sé" or ambiguous → `null`. Do not estimate from context. A blank tax rate is
  a `null`, not a typical rate.
- Build `flujos.fases` from C2/C3/D1: contributions positive, withdrawals negative, in
  today's money, net of other income (pension, rent). Spending in retirement that the
  client didn't state is a gap, not "70% of current spending".
- Likert I2, I3 and I6 are reverse-scored.
- Copy `assets/policy-default.json` to `policy.json`.

**Gate 1 — show the user and wait:** the mapped facts (a compact table), the gaps, the
assumptions that will be used (inflation, fee, any `meta_final_real` defaulted to 0), and
how you built the phases. Ask whether to fill gaps now or proceed with `[PENDIENTE]`.
If the PM has a CMA file and wants an allocation, get it here.

### 2. Calculate
```
python scripts/ips_calc.py <out>
```
Exit 1 means a shape error in your JSON: fix it and rerun. Read `calc.json`. If the
required return is unreachable (`alcanzable: false`) or very high against the CMA, that is
a finding for section 4, not something to smooth over.

If the PM gave a CMA: write `saa.json` (targets + ranges per class consistent with the
proposed tolerance and constraints), rerun `ips_calc.py`, and check `saa_check`:
`cumple_retorno` and `cumple_riesgo` must both be true, or section 7 says why not and
what the PM must choose. The return test compares the CMA return *after* fees and taxes
with the required return (which is after fees and taxes); `null` means it can't be
decided yet (missing tax rate) — say so, never read it as a pass.

### 3. Draft `ips.json`
Read `references/cfai-ips-framework.md` (what each section must hold) and
`references/revision-triggers.md` (flexibility clauses, triggers, compliance checklist).
Write sections 1–9 in professional Spanish:
- Numbers only through placeholders (`{{calc.retorno.real_pct}}`,
  `{{client.situacion.cartera_invertible}}`) or auto-tables. Never type a calculated
  number.
- Section 4: `tabla_retorno`; required vs desired; basis (real/nominal, net/gross,
  after/pre-tax); if unreachable, list the levers from C5.
- Section 5: `tabla_capacidad` and `tabla_disposicion`; then the reconciliation in words —
  which side binds, why, what it costs the client. Mention contradictions between stated
  tolerance and past behavior explicitly.
- Section 6: five `subtitulo` blocks (Horizonte, Impuestos, Liquidez, Legal y regulatorio,
  Circunstancias únicas) plus `tabla_liquidez`. Legal ends with the compliance checklist
  as `[PENDIENTE: confirmar con cumplimiento]` items.
- Section 7: `tabla_saa` when `saa_check` exists; otherwise a `[PENDIENTE]` for the PM plus
  the criteria the allocation must meet (required return, loss limit, liquidity, exclusions).
- Section 9: rebalancing method and bands, review triggers, the difference between
  rebalancing and revising the IPS, and the market-regime clause.
- `usa`: list the client facts each section relies on. Facts that genuinely don't belong
  anywhere go in `no_usados` with a reason — that list should be short.

### 4. Build
```
python scripts/build_ips_docx.py <out>
```
It validates, then writes the docx; on FAIL it lists every problem (missing section,
unresolved placeholder, unused fact) and writes nothing. Fix `ips.json` and rebuild.
Never edit the docx by hand.

### 5. Hand over (Gate 2)
Tell the user, briefly: the proposed tolerance and which side bound it, the required
return and whether it looks reachable, the open items (count and the important ones),
and the decisions left to the PM (risk conclusion, allocation, removing the draft mark).
The final version is `build_ips_docx.py <out> --final`, which only works with zero open items.

## When things don't fit
- **Institutional client** (pension, foundation, endowment, insurer, bank): out of scope
  for v1 — return objectives and risk are driven by liabilities and spending rules. Say so.
- **Couple or family**: one IPS per decision unit; record both people in
  `perfil.otras_partes` and use the household's joint figures.
- **Answers contradict each other** (e.g. B7 says stable income but B6 shows one-off
  bonuses): keep both quotes in `fuentes`, use the more conservative reading, and make
  it a `[PENDIENTE]`.
- **Updating an existing IPS**: new folder with the new date; start from the old
  `client.json`, change only what the client reported, and list the changes in section 2.
