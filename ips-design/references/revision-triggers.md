# Flexibility layer, review triggers and compliance checklist

## Why a flexibility layer
Design principle: an IPS should be flexible, not deterministic. It should capture the
client's current *and future* circumstances so the strategy adapts to different risk
environments and can take return opportunities under proper risk management. The CFA
curriculum gives the mechanics; this file turns them into IPS clauses.

Flexibility never means vagueness. Each flexible clause names *what* can move, *within
what limits*, *who* decides and *what triggers* it.

## Clauses to include
1. **Ranges, not points.** Every asset class carries a target and a permitted range.
   Drift inside the range needs no action.
2. **Revising vs rebalancing.** Rebalancing returns the portfolio to the current targets;
   revising the IPS changes the targets. Revision needs the client's signature; rebalancing
   does not.
3. **Rebalancing method.** Calendar (state frequency) or corridor (state band width per
   class). Corridors are wider for classes with higher transaction costs or lower
   volatility, and narrower for higher-volatility classes. Rebalancing costs (fees, taxes)
   count against it.
4. **Planned stage changes.** For each boundary in the multistage horizon (retirement, sale
   of business, end of education payments), state what the IPS expects to change, so the
   change is a planned review, not a surprise.
5. **Levers if the required return is out of reach.** Spend less, contribute more, extend
   the horizon, lower the bequest, accept more risk only if ability allows. Listed in the IPS
   so a bad market does not force an improvised answer.
6. **Market-regime clause.** A large drawdown or a regime change triggers a *review
   meeting*, not an automatic change of strategy. The IPS states the threshold that
   triggers the meeting (e.g. portfolio loss equal to the stated one-year loss limit).

## Review triggers (write the ones that apply)
- Calendar: at least annually (Standard III(C)).
- Life events: marriage, divorce, birth, death, inheritance, large wealth loss.
- Career: retirement, job change, move from salary to portfolio-funded spending.
- Liabilities: education funding, home purchase, planned donations.
- Tax status: change of tax residence, trust restructuring, bracket change.
- Risk tolerance: aging, unexpected wealth, a contradiction flagged at intake.
- Market: drawdown at or beyond the stated loss limit; a regime change the PM documents.

## Compliance checklist (section 6, Legal y regulatorio)
The skill does not state jurisdiction-specific rules from memory. In the legal section,
write the client's legal facts (trusts, powers, contractual limits) and then this checklist
as `[PENDIENTE: confirmar con cumplimiento]` items, adapted to the client:
- Investor profile required by the local regulator for advised services, and whether this
  IPS satisfies it or must be attached to it.
- Type of service (advised vs non-advised; discretionary vs non-discretionary management).
- Products allowed for the client's profile under local rules.
- Required disclosures (fees, conflicts of interest).
- Retention of the signed IPS and of the questionnaire.

Until the PM removes the draft mark (`build_ips_docx.py --final`), every page carries
"BORRADOR — sujeto a revisión de cumplimiento".
