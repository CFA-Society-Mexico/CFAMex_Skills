"""Render the credit-4c PDF report from recalculated values + written findings.

  python build_report.py <out_dir>

Inputs (all in <out_dir>): values.json (numbers, read back from the recalculated
workbook), findings.json (narrative), qualitative.json, checks.json, capture.csv,
policy.yaml. Numbers in charts and tables come only from values.json, so the PDF
can never disagree with the workbook.
"""
from __future__ import annotations

import argparse
import csv
import json
import sys
from pathlib import Path
from typing import Any
from xml.sax.saxutils import escape

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
import yaml  # noqa: E402
from reportlab.lib import colors  # noqa: E402
from reportlab.lib.enums import TA_LEFT  # noqa: E402
from reportlab.lib.pagesizes import A4  # noqa: E402
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet  # noqa: E402
from reportlab.lib.units import mm  # noqa: E402
from reportlab.platypus import (  # noqa: E402
    Image,
    KeepTogether,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

from registry import (  # noqa: E402
    CAPTURE_CSV,
    CHECKS_JSON,
    FINDINGS_JSON,
    POLICY_YAML,
    QUALITATIVE_JSON,
    VALUES_JSON,
    out_path,
)

# Reference palette (dataviz skill, light mode): categorical slots in fixed order + status.
SERIES = ("#2a78d6", "#eb6834", "#1baf7a")
STATUS_HEX = {"good": "#0ca30c", "warning": "#fab219", "serious": "#ec835a", "critical": "#d03b3b"}
TEXT_PRIMARY, TEXT_SECONDARY, GRID, SURFACE = "#0b0b0b", "#52514e", "#e6e5e1", "#fcfcfb"
STATUS_OF = {"Approve": "good", "OK": "good", "Pass": "good", "green": "good",
             "Review": "warning", "Tight": "warning", "amber": "warning", "WARN": "warning",
             "Reject": "critical", "Breach": "critical", "Fail": "critical", "red": "critical", "FAIL": "critical"}
CHART_W_IN, CHART_H_IN, DPI = 7.0, 2.8, 200
PAGE_W = A4[0] - 36 * mm


# ---------------------------------------------------------------- formatting
def fmt(value: Any, kind: str) -> str:
    if value is None:
        return "-"
    if isinstance(value, str):
        return value
    if kind == "x":
        return f"{value:,.2f}x"
    if kind == "pct":
        return f"{value:.1%}"
    if kind == "num":
        return f"{value:,.1f}"
    return f"{value:,.0f}"


def series(values: dict[str, Any], key: str) -> tuple[list[str], list[float | None]]:
    data = values.get(key, {}).get("values", {})
    periods = list(data)
    return periods, [v if isinstance(v, (int, float)) else None for v in data.values()]


def scalar(values: dict[str, Any], key: str) -> Any:
    return values.get(key, {}).get("value")


def para(text: str, style: ParagraphStyle) -> Paragraph:
    return Paragraph(escape(str(text)), style)


# ---------------------------------------------------------------- charts
def style_axes(ax: plt.Axes) -> None:
    ax.set_facecolor(SURFACE)
    for side in ("top", "right", "left"):
        ax.spines[side].set_visible(False)
    ax.spines["bottom"].set_color(TEXT_SECONDARY)
    ax.tick_params(colors=TEXT_SECONDARY, labelsize=8, length=0)
    ax.grid(axis="y", color=GRID, linewidth=0.8)
    ax.set_axisbelow(True)


def band_lines(ax: plt.Axes, approve: float | None, reject: float | None, kind: str) -> None:
    """Threshold lines always carry a text label: color never carries the meaning alone."""
    for value, role, word in ((approve, "good", "Approve"), (reject, "critical", "Reject")):
        if isinstance(value, (int, float)):
            ax.axhline(value, color=STATUS_HEX[role], linewidth=1.2, linestyle=(0, (4, 3)))
            ax.annotate(f"{word} {fmt(value, kind)}", xy=(1.0, value), xycoords=("axes fraction", "data"),
                        xytext=(4, 0), textcoords="offset points", va="center", fontsize=7, color=TEXT_SECONDARY)


def line_chart(path: Path, title: str, periods: list[str], lines: list[tuple[str, list[float | None]]],
               kind: str, approve: float | None = None, reject: float | None = None) -> Path:
    fig, ax = plt.subplots(figsize=(CHART_W_IN, CHART_H_IN), dpi=DPI)
    fig.patch.set_facecolor(SURFACE)
    style_axes(ax)
    for i, (label, ys) in enumerate(lines):
        xs = [p for p, y in zip(periods, ys, strict=True) if y is not None]
        yv = [y for y in ys if y is not None]
        if not yv:
            continue
        ax.plot(xs, yv, color=SERIES[i], linewidth=2, marker="o", markersize=5,
                markeredgecolor=SURFACE, markeredgewidth=1.5, label=label)
        # Direct label: latest value only, above the first series and below the second.
        ax.annotate(fmt(yv[-1], kind), xy=(xs[-1], yv[-1]),  # type: ignore[arg-type]  # categorical x
                    xytext=(0, 8 if i == 0 else -12),
                    textcoords="offset points", ha="center", fontsize=7.5, color=TEXT_PRIMARY)
    band_lines(ax, approve, reject, kind)
    ax.set_title(title, loc="left", fontsize=10, color=TEXT_PRIMARY, fontweight="bold")
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: fmt(v, kind)))
    ax.margins(x=0.06, y=0.12)
    if len(lines) > 1:
        ax.legend(frameon=False, fontsize=7.5, loc="upper center", bbox_to_anchor=(0.5, -0.12),
                  ncol=len(lines), labelcolor=TEXT_SECONDARY)
    fig.tight_layout()
    fig.savefig(path, facecolor=SURFACE)
    plt.close(fig)
    return path


def bar_chart(path: Path, title: str, categories: list[str], bars: list[tuple[str, list[float | None]]],
              kind: str, hline: tuple[str, float] | None = None, horizontal: bool = False,
              highlight_last: bool = False) -> Path:
    """highlight_last: the last category is a different quantity (e.g. value vs claims) -> slot 2 color."""
    fig, ax = plt.subplots(figsize=(CHART_W_IN, CHART_H_IN), dpi=DPI)
    fig.patch.set_facecolor(SURFACE)
    style_axes(ax)
    n = len(bars)
    width = 0.8 / max(n, 1)
    for i, (label, ys) in enumerate(bars):
        pos = [k + (i - (n - 1) / 2) * width for k in range(len(categories))]
        vals = [y if y is not None else 0 for y in ys]
        fill = [SERIES[i]] * len(vals)
        if highlight_last and n == 1:
            fill[-1] = SERIES[1]
        if horizontal:
            ax.barh(pos, vals, height=width * 0.92, color=fill, label=label)
        else:
            ax.bar(pos, vals, width=width * 0.92, color=fill, label=label)
        for x, v in zip(pos, vals, strict=True):
            if v and len(vals) <= 8 and n == 1:
                ax.annotate(fmt(v, kind), xy=(v, x) if horizontal else (x, v), xytext=(4, 0) if horizontal else (0, 3),
                            textcoords="offset points", fontsize=7, color=TEXT_PRIMARY,
                            va="center" if horizontal else "bottom", ha="left" if horizontal else "center")
    ticks = list(range(len(categories)))
    if horizontal:
        ax.set_yticks(ticks, categories)
        ax.grid(axis="x", color=GRID, linewidth=0.8)
        ax.grid(axis="y", visible=False)
        ax.xaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: fmt(v, kind)))
        ax.invert_yaxis()
    else:
        ax.set_xticks(ticks, categories)
        ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: fmt(v, kind)))
    if hline is not None:
        label, value = hline
        ax.axhline(value, color=TEXT_PRIMARY, linewidth=1.2, linestyle=(0, (4, 3)))
        ax.annotate(f"{label} {fmt(value, kind)}", xy=(1.0, value), xycoords=("axes fraction", "data"),
                    xytext=(4, 0), textcoords="offset points", va="center", fontsize=7, color=TEXT_SECONDARY)
    if not horizontal:
        ax.axhline(0, color=TEXT_SECONDARY, linewidth=0.8)
    ax.set_title(title, loc="left", fontsize=10, color=TEXT_PRIMARY, fontweight="bold")
    if n > 1:
        ax.legend(frameon=False, fontsize=7.5, loc="upper center", bbox_to_anchor=(0.5, -0.12),
                  ncol=n, labelcolor=TEXT_SECONDARY)
    fig.tight_layout()
    fig.savefig(path, facecolor=SURFACE)
    plt.close(fig)
    return path


def image(path: Path) -> Image:
    return Image(str(path), width=PAGE_W, height=PAGE_W * CHART_H_IN / CHART_W_IN)


# ---------------------------------------------------------------- tables
def status_table(rows: list[list[Any]], col_widths: list[float], status_cols: tuple[int, ...] = (),
                 styles: dict[str, ParagraphStyle] | None = None) -> Table:
    assert styles is not None
    body = [[para(c, styles["cellb"] if r == 0 else styles["cell"]) for c in row] for r, row in enumerate(rows)]
    table = Table(body, colWidths=col_widths, repeatRows=1)
    cmds: list[tuple[Any, ...]] = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F3864")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("LINEBELOW", (0, 0), (-1, -1), 0.4, colors.HexColor(GRID)),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]
    for r, row in enumerate(rows[1:], start=1):
        for c in status_cols:
            role = STATUS_OF.get(str(row[c]))
            if role:
                cmds.append(("BACKGROUND", (c, r), (c, r), colors.HexColor(STATUS_HEX[role]).clone(alpha=0.25)))
    table.setStyle(TableStyle(cmds))
    return table


def bullets(items: list[str], style: ParagraphStyle) -> list[Paragraph]:
    return [Paragraph("&bull; " + escape(str(x)), style) for x in items or []]


# ---------------------------------------------------------------- document
def make_styles() -> dict[str, ParagraphStyle]:
    base = getSampleStyleSheet()
    return {
        "title": ParagraphStyle("t", parent=base["Title"], fontSize=18, alignment=TA_LEFT, textColor=colors.HexColor(TEXT_PRIMARY)),
        "h1": ParagraphStyle("h1", parent=base["Heading1"], fontSize=14, spaceBefore=6, textColor=colors.HexColor("#1F3864")),
        "h2": ParagraphStyle("h2", parent=base["Heading2"], fontSize=11, spaceBefore=6, textColor=colors.HexColor(TEXT_PRIMARY)),
        "body": ParagraphStyle("b", parent=base["BodyText"], fontSize=9, leading=12, textColor=colors.HexColor(TEXT_PRIMARY)),
        "muted": ParagraphStyle("m", parent=base["BodyText"], fontSize=8, leading=10, textColor=colors.HexColor(TEXT_SECONDARY)),
        "bullet": ParagraphStyle("bl", parent=base["BodyText"], fontSize=9, leading=12, leftIndent=8, textColor=colors.HexColor(TEXT_PRIMARY)),
        "cell": ParagraphStyle("c", parent=base["BodyText"], fontSize=7.5, leading=9),
        "cellb": ParagraphStyle("cb", parent=base["BodyText"], fontSize=7.5, leading=9, textColor=colors.white, fontName="Helvetica-Bold"),
    }


def footer(canvas: Any, doc: Any, label: str) -> None:
    canvas.saveState()
    canvas.setFont("Helvetica", 7)
    canvas.setFillColor(colors.HexColor(TEXT_SECONDARY))
    canvas.drawString(18 * mm, 10 * mm, label)
    canvas.drawRightString(A4[0] - 18 * mm, 10 * mm, f"Page {doc.page}")
    canvas.restoreState()


def build(out_dir: Path) -> Path:
    values = json.loads(out_path(out_dir, VALUES_JSON).read_text(encoding="utf-8"))
    findings = json.loads(out_path(out_dir, FINDINGS_JSON).read_text(encoding="utf-8"))
    qual = json.loads(out_path(out_dir, QUALITATIVE_JSON).read_text(encoding="utf-8"))
    checks = json.loads(out_path(out_dir, CHECKS_JSON).read_text(encoding="utf-8"))
    policy = yaml.safe_load(out_path(out_dir, POLICY_YAML).read_text(encoding="utf-8"))
    with out_path(out_dir, CAPTURE_CSV).open(encoding="utf-8", newline="") as handle:
        docs = sorted({r["source_doc"] for r in csv.DictReader(handle)})
    meta = values["_meta"]
    iss = qual.get("issuer") or {}
    units = meta.get("units", "")
    S = make_styles()
    charts = out_dir / "charts"
    charts.mkdir(exist_ok=True)
    pm = meta["policy_metrics"]

    story: list[Any] = []
    # --- page 1: headline + scorecard
    story.append(para(f"{iss.get('name', '')} ({iss.get('ticker', '')}) - Credit analysis (4C)", S["title"]))
    story.append(para(f"As of {qual.get('as_of', '')} | Fiscal years {meta['periods'][0]}-{meta['periods'][-1]} | {units} | "
                      f"Sector: {iss.get('sector', '')}", S["muted"]))
    story.append(Spacer(1, 6))
    story.append(para(findings.get("headline", ""), S["h2"]))
    story.extend(bullets(findings.get("summary", []), S["bullet"]))
    story.append(Spacer(1, 6))
    story.append(para("Scorecard - model proposal, not a decision", S["h2"]))
    rows = [["Dimension", "Coverage", "# App.", "# Rev.", "# Rej.", "Model proposal", "Key evidence", "Analyst"]]
    for s in values.get("scorecard", []):
        rows.append([s["dimension"], s["coverage"], s["approve"], s["review"], s["reject"], s["proposal"],
                     s["evidence"], s["analyst"] or "(pending)"])
    rows.append(["Overall", "", "", "", "", scalar(values, "score.overall"), "Worst of the dimensions",
                 scalar(values, "score.analyst_overall") or "(pending)"])
    story.append(status_table(rows, [30 * mm, 18 * mm, 13 * mm, 13 * mm, 13 * mm, 22 * mm, 47 * mm, 18 * mm], (5, 7), S))
    story.append(Spacer(1, 4))
    story.append(para("The workbook proposes a status per dimension from policy bands; only core metrics can propose "
                      "Reject. The rating is the analyst's call.", S["muted"]))
    story.append(PageBreak())

    # --- Capacity
    story.append(para("1. Capacity - ability to pay", S["h1"]))
    p, de = series(values, "cap.debt_ebitda")
    _, nde = series(values, "cap.net_debt_ebitda")
    story.append(image(line_chart(charts / "leverage.png", "Leverage (adjusted, lease-inclusive)", p,
                                  [("Debt/EBITDA", de), ("Net debt/EBITDA", nde)], "x",
                                  pm["debt_ebitda"]["approve"], pm["debt_ebitda"]["reject"])))
    _, ei = series(values, "cap.ebitda_interest")
    _, bi = series(values, "cap.ebit_interest")
    story.append(image(line_chart(charts / "coverage.png", "Interest coverage", p,
                                  [("EBITDA/interest", ei), ("EBIT/interest", bi)], "x",
                                  pm["ebitda_interest"]["approve"], pm["ebitda_interest"]["reject"])))
    cap = findings.get("per_c", {}).get("capacity", {})
    story.append(para(f"Model proposal: {cap.get('proposal', '')}", S["h2"]))
    story.extend(bullets(cap.get("bullets", []), S["bullet"]))
    rows = [["Metric"] + p + ["Status", "Tier"]]
    for key, m in pm.items():
        if f"cap.{key}" not in values:
            continue
        _, ys = series(values, f"cap.{key}")
        raw = values[f"cap.{key}"]["values"]
        rows.append([m["label"]] + [fmt(raw[q], m["fmt"]) for q in p] + [scalar(values, f"cap.{key}.status"), m.get("tier", "")])
    story.append(Spacer(1, 4))
    story.append(status_table(rows, [46 * mm] + [17 * mm] * len(p) + [18 * mm, 22 * mm], (len(p) + 1,), S))
    story.append(PageBreak())

    # --- Cash flow and distributions
    story.append(para("2. Cash flow, distributions and stress", S["h1"]))
    _, focf = series(values, "adj.focf_l")
    _, div = series(values, "adj.dividends")
    _, bb = series(values, "adj.buybacks")
    dist = [(d or 0) + (b or 0) if d is not None else None for d, b in zip(div, bb, strict=True)]
    story.append(image(bar_chart(charts / "distributions.png", f"FOCF after leases vs shareholder distributions ({units})",
                                 p, [("FOCF after leases", focf), ("Dividends + buybacks", dist)], "amt")))
    _, nd = series(values, "adj.net_debt")
    story.append(image(bar_chart(charts / "net_debt.png", f"Adjusted net debt ({units})", p, [("Adjusted net debt", nd)], "amt")))
    story.append(para(f"Stress test ({meta['periods'][-1]}): EBITDA -{policy['stress']['ebitda_shock']:.0%}, "
                      f"rates +{policy['stress']['rate_shock_bp']} bp on the floating-rate share of adjusted debt "
                      f"(debt schedule; {policy['stress']['floating_share']:.0%} if no schedule). "
                      f"Verdict: {scalar(values, 'stress.verdict')}", S["h2"]))
    rows = [["Metric", "Base", "Stress", "Status under stress"]]
    for key in ("debt_ebitda", "net_debt_ebitda", "ffo_debt", "ebit_interest", "ebitda_interest", "fixed_charge_cov", "debt_service_cov"):
        m = pm[key]
        rows.append([m["label"], fmt(scalar(values, f"stress.{key}.base"), m["fmt"]),
                     fmt(scalar(values, f"stress.{key}.stress"), m["fmt"]), scalar(values, f"stress.{key}.status")])
    story.append(status_table(rows, [60 * mm, 30 * mm, 30 * mm, 40 * mm], (3,), S))
    story.extend(bullets(findings.get("stress", []), S["bullet"]))
    story.append(PageBreak())

    # --- Liquidity
    story.append(para("3. Liquidity and maturities", S["h1"]))
    sched = values.get("liq.schedule", {})
    if sched.get("has_schedule"):
        labels = ["Year 1", "Year 2", "Year 3", "Year 4", "Year 5", "After 5"]
        mats = [sched["buckets"].get(b) for b in ("y1", "y2", "y3", "y4", "y5", "after")]
        if not any(mats[1:5]):  # schedule not split by year (typical IFRS lease note): don't fake a far-dated wall
            labels, mats = ["Year 1", "After year 1 (not split by year)"], [mats[0], mats[5]]
        cash_line = (scalar(values, "liq.acc_cash") or 0) + (scalar(values, "liq.undrawn") or 0)
        story.append(image(bar_chart(charts / "maturities.png", f"Maturity wall vs cash + undrawn lines ({units})", labels,
                                     [("Maturities", mats)], "amt", hline=("Cash + undrawn", cash_line))))
    else:
        story.append(para("No debt schedule in the folder: maturity wall not available.", S["body"]))
    rows = [["Next 12 months", units]]
    for key, label in (("liq.acc_cash", "Accessible cash"), ("liq.undrawn", "Undrawn committed lines"), ("liq.ffo", "FFO"),
                       ("liq.sources", "Total sources"), ("liq.debt_due", "Debt due 12m"), ("liq.capex", "Capex"),
                       ("liq.dividends", "Dividends"), ("liq.uses", "Total uses")):
        rows.append([label, fmt(scalar(values, key), "amt")])
    rows.append(["Sources / uses", f"{fmt(scalar(values, 'liq.liquidity_su'), 'x')} ({scalar(values, 'liq.liquidity_su.status')})"])
    story.append(status_table(rows, [70 * mm, 50 * mm], (), S))
    story.extend(bullets(findings.get("liquidity", []), S["bullet"]))

    # --- Collateral
    story.append(para("4. Collateral - recovery in distress", S["h1"]))
    claims = [("Secured", "col.claims.secured"), ("Senior unsecured debt", "col.claims.senior_unsecured"),
              ("Leases", "col.claims.leases"), ("Trade payables", "col.claims.payables"), ("Subordinated", "col.claims.subordinated")]
    cats = [c for c, _ in claims] + ["Value available"]
    vals = [scalar(values, k) for _, k in claims] + [scalar(values, "col.value_available")]
    story.append(image(bar_chart(charts / "collateral.png", f"Claims vs distressed value available ({units})", cats,
                                 [("Amount", vals)], "amt", horizontal=True, highlight_last=True)))
    rows = [["Measure", "Value"],
            ["Senior unsecured recovery", f"{fmt(scalar(values, 'col.recovery_unsecured'), 'pct')} ({scalar(values, 'col.recovery_unsecured.status')})"],
            ["Asset coverage of adjusted debt", fmt(scalar(values, "col.asset_coverage"), "x")],
            ["Expected loss (PD x LGD x EAD)", fmt(scalar(values, "col.expected_loss"), "amt")]]
    story.append(status_table(rows, [70 * mm, 60 * mm], (), S))
    col_f = findings.get("per_c", {}).get("collateral", {})
    story.append(para(f"Model proposal: {col_f.get('proposal', '')}", S["h2"]))
    story.extend(bullets(col_f.get("bullets", []), S["bullet"]))
    story.append(PageBreak())

    # --- Covenants and Character
    story.append(para("5. Covenants", S["h1"]))
    covs = values.get("covenants", [])
    if covs:
        rows = [["Covenant", "Limit", "Current", "Headroom", "Status", "Stress headroom", "Cure rights"]]
        for cv in covs:
            rows.append([cv["name"], fmt(cv["limit"], "x"), fmt(cv["current"], "x"), fmt(cv["headroom"], "pct"),
                         cv["status"], fmt(cv["stress_headroom"], "pct"), cv["cure_rights"]])
        story.append(status_table(rows, [50 * mm, 16 * mm, 18 * mm, 18 * mm, 16 * mm, 24 * mm, 22 * mm], (4,), S))
    else:
        reason = ((qual.get("coverage") or {}).get("covenants") or {}).get("reason", "")
        story.append(para(f"No covenants tested. {reason}", S["body"]))
    cov_f = findings.get("per_c", {}).get("covenants", {})
    story.append(para(f"Model proposal: {cov_f.get('proposal', '')}", S["h2"]))
    story.extend(bullets(cov_f.get("bullets", []), S["bullet"]))
    story.append(para("6. Character", S["h1"]))
    rows = [["Item", "Flag", "Finding", "Source"]]
    for ch in values.get("character", []):
        rows.append([ch["item"], ch["flag"], ch["finding"] or "", ch["source"] or ""])
    story.append(status_table(rows, [34 * mm, 14 * mm, 82 * mm, 44 * mm], (1,), S))
    chr_f = findings.get("per_c", {}).get("character", {})
    story.append(para(f"Model proposal: {chr_f.get('proposal', '')}", S["h2"]))
    story.extend(bullets(chr_f.get("bullets", []), S["bullet"]))
    story.append(PageBreak())

    # --- Bear case, context, limits
    story.append(para("7. Bear case", S["h1"]))
    story.extend(bullets(findings.get("bear_case", []), S["bullet"]))
    story.append(para("8. Context", S["h1"]))
    story.extend(bullets(findings.get("context", []), S["bullet"]))
    open_doubts = [f"[{d.get('c', '')}] {d.get('doubt', '')} - {d.get('evidence', '')}"
                   for d in qual.get("doubts", []) if d.get("status") == "open"]
    if open_doubts:
        story.append(para("Open doubts", S["h2"]))
        story.extend(bullets(open_doubts, S["bullet"]))
    story.append(para("9. Method and limits", S["h1"]))
    story.extend(bullets(findings.get("limits", []), S["bullet"]))
    cov = qual.get("coverage") or {}
    rows = [["C", "Coverage", "Reason"]] + [[k.capitalize(), (cov.get(k) or {}).get("level", ""), (cov.get(k) or {}).get("reason", "")]
                                             for k in ("capacity", "collateral", "covenants", "character")]
    story.append(KeepTogether([para("Coverage of each C", S["h2"]), status_table(rows, [28 * mm, 22 * mm, 124 * mm], (), S)]))
    story.append(para(f"Integrity: workbook checks {checks.get('workbook_checks')}, formula errors {len(checks.get('errors', []))}, "
                      f"data gaps (#N/A) {len(checks.get('gaps', []))}. Recalculated with {checks.get('engine')}.", S["body"]))
    story.append(para(f"Policy: {policy.get('name', '')}. Bands marked 'skill default' are not the analyst's policy until adopted.", S["body"]))
    story.append(para("Framework: CFA L1/L2 credit analysis (4 Cs, EL = PD x LGD x EAD); S&P corporate ratios and adjustments "
                      "(2019 criteria: adjusted debt incl. leases, FFO, FOCF, DCF). Every hard-coded figure is cited in the "
                      "workbook's Sources sheet.", S["body"]))
    story.append(para("Source documents", S["h2"]))
    story.append(para("; ".join(docs), S["muted"]))

    pdf = out_dir / f"{meta['ticker']}_credit-4c_{qual.get('as_of', '')}.pdf"
    label = f"{iss.get('ticker', '')} credit-4c | model proposal, not a rating | {qual.get('as_of', '')}"
    doc = SimpleDocTemplate(str(pdf), pagesize=A4, leftMargin=18 * mm, rightMargin=18 * mm, topMargin=16 * mm,
                            bottomMargin=16 * mm, title=f"{iss.get('ticker', '')} credit analysis (4C)")
    doc.build(story, onFirstPage=lambda c, d: footer(c, d, label), onLaterPages=lambda c, d: footer(c, d, label))
    return pdf


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description="Render the credit-4c PDF report.")
    parser.add_argument("out", type=Path)
    args = parser.parse_args(argv)
    pdf = build(args.out)
    print(f"[ok] report -> {str(pdf).encode('ascii', 'replace').decode('ascii')}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
