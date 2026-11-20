# Risk tolerance — ability vs willingness

`scripts/ips_calc.py` scores both sides from `client.json` using the thresholds in
`policy.json` (copied from `assets/policy-default.json`). Your job is to read the scores,
check them against the client's own words, and write the conclusion with evidence. The
script *proposes*; the PM decides (the `Decisión del PM` line in section 5 stays empty).

## Why the scoring is split into factors
LLM advisers collapse multi-factor suitability into one input — usually the self-reported
risk tolerance (Ross & Lo, MIT 2026: risk tolerance carried 57–88% of allocation weight;
income, horizon and liquidity barely moved the answer). The defense is structural:
- Ability is scored factor by factor from hard facts, never from the questionnaire mood.
- Willingness is scored from three independent sources (questionnaire, loss scenario,
  past behavior) and contradictions are flagged, not averaged away.
- Section 5 of the IPS must show every ability factor with its value. A risk section that
  only cites the questionnaire is the failure this skill exists to prevent.

## Ability (capacidad) — factors
| Factor | Computed from | Core? |
|---|---|---|
| `horizonte` | total horizon in years (sum of phases) | yes |
| `tasa_retiro` | annual net withdrawal from portfolio / investable portfolio (peak withdrawal phase) | yes |
| `liquidez` | (emergency reserve shortfall + one-time needs within 24 months) / portfolio | yes |
| `retorno_requerido` | required real return (large goals relative to wealth reduce ability) | no |
| `estabilidad_ingreso` | client answer: alta / media / baja | no |
| `deuda` | total debt / net worth | no |
| `flexibilidad` | count of "sí" to: can reduce spending; can delay retirement or goal; has other income or assets outside the portfolio | no |

Each factor scores `bajo` / `medio` / `alto` against `policy.json` bands. Aggregation:
1. Overall = median of the available factor scores (ties round down).
2. Cap: one core factor `bajo` caps the overall at `medio`; two or more cap it at `bajo`.
3. More than `max_missing_factors` factors missing → `indeterminado`, and the risk
   conclusion becomes `[PENDIENTE]`.

## Willingness (disposición) — sources
| Source | Mapping |
|---|---|
| Questionnaire (6 Likert items, 1–5; items marked `reverse` are inverted) | mean < 2.5 `bajo`; 2.5–3.5 `medio`; > 3.5 `alto` |
| Loss scenario: max one-year loss the client would hold through | < 10% `bajo`; 10–20% `medio`; > 20% `alto` |
| Past behavior in the last large drawdown | vendió `bajo`; mantuvo `medio`; compró más `alto`; sin experiencia → not scored |

Overall willingness = median of available sources (ties round down). If any two sources
differ by two levels (`bajo` vs `alto`), the script sets `contradiccion: true`; write it up
in section 5 and list it for the PM — stated tolerance is noisy and moves with the market cycle.

## Conclusion
- **Proposed tolerance = min(ability, willingness).**
- Ability > willingness: tolerance follows willingness. Say what that costs (lower expected
  return, maybe a required return that becomes hard to reach).
- Willingness > ability: tolerance follows ability. Say that the client was told why.
- Quantitative limit: the maximum one-year loss = the lower of the client's stated loss and
  the cap for the proposed tolerance level in `policy.json` (`limite_perdida_por_nivel`).

## Personality type (L3 R8)
Assign one with evidence; it shapes how the IPS is communicated, not the numbers.
| Type | Signals |
|---|---|
| Cauteloso | Risk averse, feeling-based; avoids decisions; dislikes volatility; low turnover |
| Metódico | Risk averse, thinking-based; asks for data and research; conservative |
| Individualista | Less averse, thinking-based; confident in own analysis; contrarian |
| Espontáneo | Less averse, feeling-based; follows trends; wants frequent changes |

## Stage of life
Foundation (fundación), accumulation (acumulación), maintenance (mantenimiento),
distribution (distribución). Progression is not linear: windfalls or setbacks move a client
between stages regardless of age.

## Source of wealth
- Active (built a business): suggests willingness, but business risk ≠ investment risk;
  check.
- Passive (inheritance, windfall): often less experience and a wish to preserve.
- Saved from employment: usually a preference for security.
