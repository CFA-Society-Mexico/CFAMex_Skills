"""Deterministic calculations for an IPS draft.

Reads <out>/client.json and <out>/policy.json (plus optional cma.json and saa.json),
writes <out>/calc.json. The LLM never computes these numbers; it quotes them.

    python ips_calc.py <out_dir>

Exit code 1 on invalid input (types, missing files); 0 otherwise. Data gaps are not
errors: they become entries in calc.json["pendientes"].
Console output is ASCII only.
"""
from __future__ import annotations

import json
import math
import statistics
import sys
from collections.abc import Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Literal

Level = Literal["bajo", "medio", "alto"]
LEVELS: tuple[Level, ...] = ("bajo", "medio", "alto")
RANK: dict[str, int] = {lvl: i for i, lvl in enumerate(LEVELS)}

RATE_LO = -0.5
RATE_HI = 1.0


class InputError(ValueError):
    """Raised when client.json / policy.json has the wrong shape."""


@dataclass(frozen=True)
class Phase:
    nombre: str
    anios: int
    flujo_anual_real: float


# ---------------------------------------------------------------- helpers

def fmt_pct(x: float | None) -> str:
    return "n.d." if x is None else f"{x * 100:.2f}%"


def fmt_amt(x: float | None) -> str:
    return "n.d." if x is None else f"{x:,.0f}"


def num(value: Any, path: str) -> float | None:
    """Return a float, None for null; raise on non-numeric."""
    if value is None:
        return None
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise InputError(f"{path} must be a number or null, got {value!r}")
    if isinstance(value, float) and math.isnan(value):
        return None
    return float(value)


def get(d: dict[str, Any], dotted: str) -> Any:
    cur: Any = d
    for key in dotted.split("."):
        if not isinstance(cur, dict) or key not in cur:
            return None
        cur = cur[key]
    return cur


def band_low_good(value: float, alto_hasta: float, bajo_mayor_a: float) -> Level:
    """Smaller is better (withdrawal rate, liquidity drain, debt)."""
    if value <= alto_hasta:
        return "alto"
    if value > bajo_mayor_a:
        return "bajo"
    return "medio"


def median_level(levels: Sequence[Level]) -> Level:
    """Median of ordinal levels; ties round down (conservative)."""
    ranks = sorted(RANK[lv] for lv in levels)
    return LEVELS[int(math.floor(statistics.median_low(ranks)))]


def min_level(a: Level, b: Level) -> Level:
    return a if RANK[a] <= RANK[b] else b


# ---------------------------------------------------------------- returns

def parse_phases(client: dict[str, Any]) -> list[Phase]:
    raw = get(client, "flujos.fases")
    if raw is None:
        return []
    if not isinstance(raw, list):
        raise InputError("flujos.fases must be a list")
    phases: list[Phase] = []
    for i, p in enumerate(raw):
        anios = p.get("anios")
        flujo = num(p.get("flujo_anual_real"), f"flujos.fases[{i}].flujo_anual_real")
        if not isinstance(anios, int) or anios <= 0:
            raise InputError(f"flujos.fases[{i}].anios must be a positive integer")
        if flujo is None:
            raise InputError(f"flujos.fases[{i}].flujo_anual_real is null; drop the phase or fill it")
        phases.append(Phase(str(p.get("nombre", f"Fase {i + 1}")), anios, flujo))
    return phases


def yearly_flows(phases: Sequence[Phase]) -> list[float]:
    flows: list[float] = []
    for ph in phases:
        flows.extend([ph.flujo_anual_real] * ph.anios)
    return flows


def terminal_wealth(w0: float, flows: Sequence[float], r: float) -> float:
    """Wealth after len(flows) years; flows at end of each year."""
    w = w0
    for f in flows:
        w = w * (1.0 + r) + f
    return w


def wealth_path(w0: float, flows: Sequence[float], r: float) -> list[float]:
    path = [w0]
    w = w0
    for f in flows:
        w = w * (1.0 + r) + f
        path.append(w)
    return path


def solve_required_return(w0: float, flows: Sequence[float], target: float) -> float | None:
    """Rate r such that terminal wealth == target, by bisection. None if no root in range."""
    if not flows:
        return None

    def gap(r: float) -> float:
        return terminal_wealth(w0, flows, r) - target

    lo, hi = RATE_LO, RATE_HI
    g_lo, g_hi = gap(lo), gap(hi)
    if g_lo > 0:
        # Even a -50% annual return reaches the target: required return below range.
        return RATE_LO
    if g_hi < 0:
        return None
    for _ in range(200):
        mid = (lo + hi) / 2.0
        if gap(mid) < 0:
            lo = mid
        else:
            hi = mid
        if hi - lo < 1e-10:
            break
    return (lo + hi) / 2.0


def return_block(
    w0: float | None,
    phases: Sequence[Phase],
    target: float | None,
    inflation: float | None,
    fee: float | None,
    tax: float | None,
    pend: list[str],
    label: str,
) -> dict[str, Any] | None:
    if w0 is None or not phases:
        pend.append(f"{label}: falta cartera invertible o fases de flujo; no se puede calcular.")
        return None
    notes: list[str] = []
    if target is None:
        target = 0.0
        notes.append("Meta final no declarada: se asume 0 (agotar la cartera al final). Confirmar.")
        pend.append(f"{label}: confirmar patrimonio deseado al final del horizonte.")
    flows = yearly_flows(phases)
    real = solve_required_return(w0, flows, target)
    reachable = real is not None
    if not reachable:
        notes.append(f"No existe retorno entre {fmt_pct(RATE_LO)} y {fmt_pct(RATE_HI)} que cumpla la meta.")
    nominal = None if real is None or inflation is None else (1 + real) * (1 + inflation) - 1
    gross = None if nominal is None or fee is None else nominal + fee
    if nominal is not None and fee is None:
        pend.append(f"{label}: comisión anual no definida; retorno reportado neto de comisiones.")
    pretax = None if gross is None or tax is None else gross / (1 - tax)
    if gross is not None and tax is None:
        pend.append(f"{label}: tasa de impuestos sobre rendimientos no proporcionada; falta retorno antes de impuestos.")
    return {
        "real_pct": real,
        "nominal_pct": nominal,
        "bruto_comisiones_pct": gross,
        "antes_impuestos_pct": pretax,
        "real_fmt": fmt_pct(real),
        "nominal_fmt": fmt_pct(nominal),
        "bruto_comisiones_fmt": fmt_pct(gross),
        "antes_impuestos_fmt": fmt_pct(pretax),
        "horizonte_anios": len(flows),
        "meta_final_real": target,
        "meta_final_fmt": fmt_amt(target),
        "cartera_inicial_fmt": fmt_amt(w0),
        "alcanzable": reachable,
        "metodo": (
            "Tasa que iguala la cartera inicial más aportaciones menos retiros (en dinero de hoy) "
            "con la meta final; nominal = (1 + real)(1 + inflación) − 1; bruto = nominal + comisión; "
            "antes de impuestos = bruto / (1 - tasa)."
        ),
        "notas": notes,
    }


# ---------------------------------------------------------------- liquidity

def liquidity_block(client: dict[str, Any], policy: dict[str, Any], pend: list[str]) -> dict[str, Any]:
    gasto = num(get(client, "situacion.gasto_anual"), "situacion.gasto_anual")
    meses = num(get(client, "liquidez.meses_reserva"), "liquidez.meses_reserva")
    actual = num(get(client, "liquidez.reserva_actual"), "liquidez.reserva_actual")
    w0 = num(get(client, "situacion.cartera_invertible"), "situacion.cartera_invertible")
    horizon_m = int(policy.get("liquidez_horizonte_meses", 24))
    needs_raw = get(client, "liquidez.necesidades_unicas") or []
    near = 0.0
    for i, n in enumerate(needs_raw):
        monto = num(n.get("monto"), f"liquidez.necesidades_unicas[{i}].monto")
        m = n.get("meses")
        if monto is None or m is None:
            pend.append(f"Liquidez: necesidad '{n.get('descripcion')}' sin monto o fecha.")
            continue
        if m <= horizon_m:
            near += monto
    required = None if gasto is None or meses is None else gasto / 12.0 * meses
    if required is None:
        pend.append("Liquidez: falta gasto anual o meses de reserva deseados.")
    shortfall = None if required is None or actual is None else max(0.0, required - actual)
    if required is not None and actual is None:
        pend.append("Liquidez: falta reserva actual disponible.")
    ratio = None
    if w0 and shortfall is not None:
        ratio = (shortfall + near) / w0
    return {
        "reserva_requerida": required,
        "reserva_actual": actual,
        "faltante_reserva": shortfall,
        "necesidades_horizonte": near,
        "horizonte_meses": horizon_m,
        "ratio_pct": ratio,
        "reserva_requerida_fmt": fmt_amt(required),
        "reserva_actual_fmt": fmt_amt(actual),
        "faltante_reserva_fmt": fmt_amt(shortfall),
        "necesidades_horizonte_fmt": fmt_amt(near),
        "ratio_fmt": fmt_pct(ratio),
    }


# ---------------------------------------------------------------- ability

def peak_withdrawal_rate(w0: float | None, phases: Sequence[Phase], r: float | None) -> float | None:
    """Largest withdrawal / wealth at the start of a withdrawal phase (wealth projected at r)."""
    if w0 is None or not phases:
        return None
    rate = r if r is not None else 0.0
    flows = yearly_flows(phases)
    path = wealth_path(w0, flows, rate)
    start = 0
    peak = 0.0
    for ph in phases:
        w_start = path[start]
        if ph.flujo_anual_real < 0:
            if w_start <= 0:
                return math.inf
            peak = max(peak, -ph.flujo_anual_real / w_start)
        start += ph.anios
    return peak


def ability_block(
    client: dict[str, Any],
    policy: dict[str, Any],
    phases: Sequence[Phase],
    real: float | None,
    liq_ratio: float | None,
    pend: list[str],
) -> dict[str, Any]:
    cap = policy["capacidad"]
    w0 = num(get(client, "situacion.cartera_invertible"), "situacion.cartera_invertible")
    factors: list[dict[str, Any]] = []

    def add(fid: str, label: str, value: float | None, fmt: str, level: Level | None, core: bool) -> None:
        factors.append({"id": fid, "etiqueta": label, "valor": value, "valor_fmt": fmt,
                        "nivel": level, "core": core})

    horizon = float(sum(p.anios for p in phases)) if phases else None
    h = cap["horizonte_anios"]
    add("horizonte", "Horizonte total", horizon,
        "n.d." if horizon is None else f"{horizon:.0f} años",
        None if horizon is None else ("bajo" if horizon < h["bajo_menor_a"] else "alto" if horizon > h["alto_mayor_a"] else "medio"),
        True)

    wr = peak_withdrawal_rate(w0, phases, real)
    t = cap["tasa_retiro"]
    add("tasa_retiro", "Tasa de retiro máxima sobre la cartera", wr,
        "n.d." if wr is None else ("sin cartera" if math.isinf(wr) else fmt_pct(wr)),
        None if wr is None else band_low_good(wr, t["alto_hasta"], t["bajo_mayor_a"]), True)

    lq = cap["liquidez_ratio"]
    add("liquidez", "Necesidades de liquidez / cartera", liq_ratio, fmt_pct(liq_ratio),
        None if liq_ratio is None else band_low_good(liq_ratio, lq["alto_hasta"], lq["bajo_mayor_a"]), True)

    rr = cap["retorno_requerido_real"]
    add("retorno_requerido", "Retorno real requerido (tamaño de metas vs riqueza)", real, fmt_pct(real),
        None if real is None else band_low_good(real, rr["alto_hasta"], rr["bajo_mayor_a"]), False)

    est = get(client, "situacion.estabilidad_ingreso")
    est_level: Level | None = {"alta": "alto", "media": "medio", "baja": "bajo"}.get(str(est)) if est else None  # type: ignore[assignment]
    add("estabilidad_ingreso", "Estabilidad del ingreso", None, str(est) if est else "n.d.", est_level, False)

    debt = num(get(client, "situacion.deuda_total"), "situacion.deuda_total")
    nw = num(get(client, "situacion.patrimonio_neto"), "situacion.patrimonio_neto")
    dr = None if debt is None or not nw or nw <= 0 else debt / nw
    d = cap["deuda_patrimonio"]
    add("deuda", "Deuda / patrimonio neto", dr, fmt_pct(dr),
        None if dr is None else band_low_good(dr, d["alto_hasta"], d["bajo_mayor_a"]), False)

    flex_keys = ("puede_reducir_gasto", "puede_retrasar_meta", "tiene_otros_ingresos")
    flex_vals = [get(client, f"flexibilidad.{k}") for k in flex_keys]
    known = [v for v in flex_vals if isinstance(v, bool)]
    f = cap["flexibilidad_si"]
    if known:
        yes = sum(1 for v in known if v)
        flev: Level = "alto" if yes >= f["alto_desde"] else "bajo" if yes <= f["bajo_hasta"] else "medio"
        add("flexibilidad", "Flexibilidad (ajustar gasto, plazo u otros ingresos)", float(yes),
            f"{yes} de {len(known)} respuestas «sí»", flev, False)
    else:
        add("flexibilidad", "Flexibilidad (ajustar gasto, plazo u otros ingresos)", None, "n.d.", None, False)

    scored = [x for x in factors if x["nivel"] is not None]
    missing = [x["etiqueta"] for x in factors if x["nivel"] is None]
    max_missing = int(cap.get("max_factores_faltantes", 2))
    if len(missing) > max_missing or not scored:
        pend.append("Capacidad de riesgo indeterminada: faltan " + ", ".join(missing) + ".")
        return {"factores": factors, "mediana": None, "tope": None, "nivel": "indeterminado",
                "faltantes": missing}
    med = median_level([x["nivel"] for x in scored])
    core_low = sum(1 for x in scored if x["core"] and x["nivel"] == "bajo")
    cap_level: Level | None = "bajo" if core_low >= 2 else "medio" if core_low == 1 else None
    final = med if cap_level is None else min_level(med, cap_level)
    for m in missing:
        pend.append(f"Capacidad: factor sin dato ({m}).")
    return {"factores": factors, "mediana": med, "tope": cap_level, "nivel": final,
            "faltantes": missing}


# ---------------------------------------------------------------- willingness

def willingness_block(client: dict[str, Any], policy: dict[str, Any], pend: list[str]) -> dict[str, Any]:
    pol = policy["disposicion"]
    sources: list[dict[str, Any]] = []
    items = get(client, "disposicion.likert") or []
    scores: list[float] = []
    for i, it in enumerate(items):
        v = num(it.get("valor"), f"disposicion.likert[{i}].valor")
        if v is None:
            continue
        if not 1 <= v <= 5:
            raise InputError(f"disposicion.likert[{i}].valor must be 1-5")
        scores.append(6 - v if it.get("reverse") else v)
    if scores:
        mean = sum(scores) / len(scores)
        lv: Level = "bajo" if mean < pol["likert_medio_desde"] else "alto" if mean > pol["likert_alto_mayor_a"] else "medio"
        sources.append({"id": "cuestionario", "etiqueta": "Cuestionario (promedio 1-5)",
                        "valor": mean, "valor_fmt": f"{mean:.2f} ({len(scores)} reactivos)", "nivel": lv})
    else:
        pend.append("Disposición: cuestionario de actitud ante el riesgo sin respuestas.")

    loss = num(get(client, "disposicion.perdida_maxima_tolerada"), "disposicion.perdida_maxima_tolerada")
    if loss is not None:
        llv: Level = "bajo" if loss < pol["perdida_medio_desde"] else "alto" if loss > pol["perdida_alto_mayor_a"] else "medio"
        sources.append({"id": "perdida", "etiqueta": "Pérdida máxima tolerada en un año",
                        "valor": loss, "valor_fmt": fmt_pct(loss), "nivel": llv})
    else:
        pend.append("Disposición: falta la pérdida máxima tolerada en un año.")

    beh = get(client, "disposicion.comportamiento_pasado")
    beh_map: dict[str, Level] = {"vendio": "bajo", "mantuvo": "medio", "compro": "alto"}
    if beh in beh_map:
        sources.append({"id": "comportamiento", "etiqueta": "Comportamiento en la última caída fuerte",
                        "valor": None, "valor_fmt": str(beh), "nivel": beh_map[str(beh)]})

    if not sources:
        return {"fuentes": [], "nivel": "indeterminado", "contradiccion": False}
    levels = [s["nivel"] for s in sources]
    contradiction = max(RANK[x] for x in levels) - min(RANK[x] for x in levels) >= 2
    if contradiction:
        pend.append("Disposición: fuentes contradictorias (diferencia de dos niveles); validar con el cliente.")
    return {"fuentes": sources, "nivel": median_level(levels), "contradiccion": contradiction}


# ---------------------------------------------------------------- SAA check

def saa_block(out: Path, policy: dict[str, Any], required_nominal: float | None,
              loss_limit: float | None, fee: float | None, tax: float | None,
              pend: list[str]) -> dict[str, Any] | None:
    """Portfolio E(R), sigma and shortfall from the CMA, checked against the IPS.

    The required nominal return is the rate the portfolio must compound at after fees
    and taxes, while CMA returns are gross. The return test therefore compares
    (E(R) - fee) * (1 - tax) against it. Without a tax rate the test cannot be
    decided (None), never passed on a pre-tax figure.
    """
    cma_p, saa_p = out / "cma.json", out / "saa.json"
    if not (cma_p.exists() and saa_p.exists()):
        return None
    cma = json.loads(cma_p.read_text(encoding="utf-8"))
    saa = json.loads(saa_p.read_text(encoding="utf-8"))
    names = [c["clase"] for c in cma["clases"]]
    er = [num(c["retorno_esperado"], f"cma.{c['clase']}.retorno_esperado") for c in cma["clases"]]
    vol = [num(c["volatilidad"], f"cma.{c['clase']}.volatilidad") for c in cma["clases"]]
    corr = cma["correlaciones"]
    if any(x is None for x in er + vol) or any(v is None for row in corr for v in row):
        raise InputError("cma.json has null expected returns, volatilities or correlations")
    er_f: list[float] = [float(x) for x in er if x is not None]
    vol_f: list[float] = [float(x) for x in vol if x is not None]
    n = len(names)
    for i in range(n):
        for j in range(n):
            if abs(corr[i][j] - corr[j][i]) > 1e-9:
                raise InputError("cma.json correlation matrix is not symmetric")
    weights = [0.0] * n
    rows: list[dict[str, Any]] = []
    for c in saa["clases"]:
        if c["clase"] not in names:
            raise InputError(f"saa.json class '{c['clase']}' not in cma.json")
        if not c["min"] <= c["objetivo"] <= c["max"]:
            raise InputError(f"saa.json class '{c['clase']}': objetivo outside [min, max]")
        weights[names.index(c["clase"])] = float(c["objetivo"])
        rows.append({**c, "objetivo_fmt": fmt_pct(c["objetivo"]),
                     "rango_fmt": f"{fmt_pct(c['min'])} - {fmt_pct(c['max'])}"})
    total = sum(weights)
    if abs(total - 1.0) > 1e-6:
        raise InputError(f"saa.json targets sum to {total:.4f}, not 1")
    e = sum(w * r for w, r in zip(weights, er_f, strict=True))
    var = sum(weights[i] * weights[j] * vol_f[i] * vol_f[j] * corr[i][j]
              for i in range(n) for j in range(n))
    sd = math.sqrt(var)
    z = float(policy.get("shortfall_z", 2.0))
    e_net = e - (fee or 0.0)
    if fee is None:
        pend.append("Asignación: comisión no definida; la verificación de retorno asume comisión 0 (optimista).")
    e_after_tax = None if tax is None else e_net * (1.0 - tax)
    if tax is None:
        pend.append("Asignación: sin tasa de impuestos sobre rendimientos no se puede verificar "
                    "si la mezcla alcanza el retorno requerido (que es después de impuestos).")
    shortfall = e_net - z * sd
    meets_return = (None if required_nominal is None or e_after_tax is None
                    else e_after_tax >= required_nominal)
    return {
        "cma_fecha": cma.get("fecha"), "cma_fuente": cma.get("fuente"),
        "clases": rows,
        "retorno_esperado_pct": e, "retorno_neto_pct": e_net,
        "retorno_despues_impuestos_pct": e_after_tax, "volatilidad_pct": sd,
        "shortfall_pct": shortfall,
        "retorno_requerido_nominal_pct": required_nominal,
        "retorno_esperado_fmt": fmt_pct(e), "retorno_neto_fmt": fmt_pct(e_net),
        "retorno_despues_impuestos_fmt": fmt_pct(e_after_tax),
        "retorno_requerido_nominal_fmt": fmt_pct(required_nominal),
        "volatilidad_fmt": fmt_pct(sd), "shortfall_fmt": fmt_pct(shortfall),
        "cumple_retorno": meets_return,
        "cumple_riesgo": None if loss_limit is None else shortfall >= -loss_limit,
    }


# ---------------------------------------------------------------- main

def assumption(policy: dict[str, Any], client: dict[str, Any], key: str) -> tuple[float | None, str]:
    override = num(get(client, f"supuestos_cliente.{key}"), f"supuestos_cliente.{key}")
    if override is not None:
        return override, "Proporcionado por el cliente o el PM"
    entry = policy["supuestos"].get(key, {})
    return num(entry.get("valor"), f"policy.supuestos.{key}"), str(entry.get("fuente", ""))


def compute(out: Path) -> dict[str, Any]:
    client = json.loads((out / "client.json").read_text(encoding="utf-8"))
    policy = json.loads((out / "policy.json").read_text(encoding="utf-8"))
    pend: list[str] = []

    inflation, inf_src = assumption(policy, client, "inflacion")
    fee, fee_src = assumption(policy, client, "comision_anual")
    tax = num(get(client, "impuestos.tasa_sobre_rendimientos"), "impuestos.tasa_sobre_rendimientos")
    if tax is not None and not 0 <= tax < 1:
        raise InputError("impuestos.tasa_sobre_rendimientos must be in [0, 1)")

    w0 = num(get(client, "situacion.cartera_invertible"), "situacion.cartera_invertible")
    phases = parse_phases(client)
    target = num(get(client, "flujos.meta_final_real"), "flujos.meta_final_real")
    ret = return_block(w0, phases, target, inflation, fee, tax, pend, "Retorno requerido")

    desired = None
    d_target = num(get(client, "deseado.meta_final_real"), "deseado.meta_final_real")
    if d_target is not None:
        desired = return_block(w0, phases, d_target, inflation, fee, tax, [], "Retorno deseado")

    liq = liquidity_block(client, policy, pend)
    real = None if ret is None else ret["real_pct"]
    ability = ability_block(client, policy, phases, real, liq["ratio_pct"], pend)
    will = willingness_block(client, policy, pend)

    if "indeterminado" in (ability["nivel"], will["nivel"]):
        proposal: str = "indeterminado"
        limit = None
    else:
        proposal = min_level(ability["nivel"], will["nivel"])
        cap_limit = float(policy["limite_perdida_por_nivel"][proposal])
        stated = num(get(client, "disposicion.perdida_maxima_tolerada"), "disposicion.perdida_maxima_tolerada")
        limit = cap_limit if stated is None else min(cap_limit, stated)

    path = None
    if w0 is not None and phases:
        flows = yearly_flows(phases)
        rr = real if real is not None else 0.0
        full = wealth_path(w0, flows, rr)
        path, start = [], 0
        for ph in phases:
            path.append({"fase": ph.nombre, "inicio_anio": start, "anios": ph.anios,
                         "flujo_anual_real_fmt": fmt_amt(ph.flujo_anual_real),
                         "riqueza_inicio_fmt": fmt_amt(full[start])})
            start += ph.anios

    nominal = None if ret is None else ret["nominal_pct"]
    saa = saa_block(out, policy, nominal, limit, fee, tax, pend)
    if saa is None:
        pend.append("Asignación estratégica: sin CMA/SAA; la define el PM.")

    return {
        "as_of": get(client, "meta.fecha_ips"),
        "supuestos": {
            "inflacion": {"valor": inflation, "valor_fmt": fmt_pct(inflation), "fuente": inf_src},
            "comision_anual": {"valor": fee, "valor_fmt": fmt_pct(fee), "fuente": fee_src},
            "tasa_sobre_rendimientos": {"valor": tax, "valor_fmt": fmt_pct(tax),
                                        "fuente": "Cliente / contador" if tax is not None else "No proporcionada"},
        },
        "retorno": ret,
        "retorno_deseado": desired,
        "liquidez": liq,
        "capacidad": ability,
        "disposicion": will,
        "tolerancia": {"propuesta": proposal, "limite_perdida_pct": limit,
                       "limite_perdida_fmt": fmt_pct(limit)},
        "trayectoria": path,
        "saa_check": saa,
        "pendientes": pend,
    }


def main(argv: Sequence[str]) -> int:
    if len(argv) != 2:
        print("usage: python ips_calc.py <out_dir>")
        return 2
    out = Path(argv[1])
    try:
        calc = compute(out)
    except (InputError, FileNotFoundError, json.JSONDecodeError, KeyError) as exc:
        print(f"[x] FAIL: {exc}")
        return 1
    (out / "calc.json").write_text(json.dumps(calc, ensure_ascii=False, indent=2), encoding="utf-8")
    r = calc["retorno"]
    print(f"[ok] calc.json written to {out}")
    if r:
        print(f"     required real return: {r['real_fmt']}  nominal: {r['nominal_fmt']}")
    print(f"     ability: {calc['capacidad']['nivel']}  willingness: {calc['disposicion']['nivel']}"
          f"  proposal: {calc['tolerancia']['propuesta']}")
    print(f"     open items: {len(calc['pendientes'])}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
