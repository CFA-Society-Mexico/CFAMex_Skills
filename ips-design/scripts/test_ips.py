"""pytest suite for ips_calc.py and build_ips_docx.py.

    python -m pytest -q test_ips.py
"""
from __future__ import annotations

import json
import shutil
from pathlib import Path
from typing import Any

import pytest

import build_ips_docx as builder
import ips_calc as calc

POLICY = Path(__file__).resolve().parent.parent / "assets" / "policy-default.json"


def base_client() -> dict[str, Any]:
    return {
        "meta": {"cliente": "Cliente Prueba", "slug": "cliente-prueba", "fecha_ips": "2026-11-19",
                 "fecha_datos": "2026-10-31", "moneda": "MXN", "asesor": "PM", "tipo_inversionista": "individual"},
        "perfil": {"edad": 40, "edad_retiro": 65, "fuente_riqueza": "ahorro_empleo"},
        "situacion": {"cartera_invertible": 1_000_000, "patrimonio_neto": 2_000_000, "deuda_total": 100_000,
                      "ingreso_anual": 800_000, "estabilidad_ingreso": "alta", "gasto_anual": 480_000},
        "flujos": {"fases": [{"nombre": "Acumulación", "anios": 25, "flujo_anual_real": 100_000},
                             {"nombre": "Retiro", "anios": 30, "flujo_anual_real": -300_000}],
                   "meta_final_real": 0},
        "liquidez": {"meses_reserva": 6, "reserva_actual": 240_000, "necesidades_unicas": []},
        "impuestos": {"residencia_fiscal": "México", "tasa_sobre_rendimientos": None},
        "unicas": ["Excluir tabaco"],
        "disposicion": {"likert": [{"id": "I1", "valor": 4, "reverse": False},
                                   {"id": "I2", "valor": 2, "reverse": True}],
                        "perdida_maxima_tolerada": 0.15, "comportamiento_pasado": "mantuvo"},
        "flexibilidad": {"puede_reducir_gasto": True, "puede_retrasar_meta": True, "tiene_otros_ingresos": False},
        "fuentes": {},
    }


@pytest.fixture
def out_dir(tmp_path: Path) -> Path:
    shutil.copy(POLICY, tmp_path / "policy.json")
    return tmp_path


def write(out: Path, client: dict[str, Any]) -> dict[str, Any]:
    (out / "client.json").write_text(json.dumps(client, ensure_ascii=False), encoding="utf-8")
    result = calc.compute(out)
    (out / "calc.json").write_text(json.dumps(result, ensure_ascii=False), encoding="utf-8")
    return result


# ---------------------------------------------------------------- returns

def test_required_return_pure_growth() -> None:
    r = calc.solve_required_return(100.0, [0.0] * 10, 200.0)
    assert r == pytest.approx(2 ** 0.1 - 1, abs=1e-8)


def test_required_return_annuity() -> None:
    # PV of 30 withdrawals of 10 at 5% -> required return is 5%.
    pv = sum(10 / 1.05 ** t for t in range(1, 31))
    r = calc.solve_required_return(pv, [-10.0] * 30, 0.0)
    assert r == pytest.approx(0.05, abs=1e-8)


def test_unreachable_goal_returns_none() -> None:
    assert calc.solve_required_return(1.0, [-1000.0] * 5, 0.0) is None


def test_nominal_and_pretax_chain(out_dir: Path) -> None:
    c = base_client()
    c["impuestos"]["tasa_sobre_rendimientos"] = 0.10
    c["supuestos_cliente"] = {"comision_anual": 0.01}
    res = write(out_dir, c)
    r = res["retorno"]
    assert r["nominal_pct"] == pytest.approx((1 + r["real_pct"]) * 1.03 - 1)
    assert r["bruto_comisiones_pct"] == pytest.approx(r["nominal_pct"] + 0.01)
    assert r["antes_impuestos_pct"] == pytest.approx(r["bruto_comisiones_pct"] / 0.9)


def test_missing_target_is_assumption_and_pending(out_dir: Path) -> None:
    c = base_client()
    c["flujos"]["meta_final_real"] = None
    res = write(out_dir, c)
    assert res["retorno"]["meta_final_real"] == 0
    assert any("patrimonio deseado" in p for p in res["pendientes"])


def test_missing_tax_never_invents_pretax(out_dir: Path) -> None:
    c = base_client()
    c["supuestos_cliente"] = {"comision_anual": 0.01}
    res = write(out_dir, c)
    assert res["retorno"]["antes_impuestos_pct"] is None
    assert any("impuestos" in p for p in res["pendientes"])


# ---------------------------------------------------------------- risk

def test_median_ties_round_down() -> None:
    assert calc.median_level(["bajo", "alto"]) == "bajo"
    assert calc.median_level(["medio", "alto", "alto"]) == "alto"


def test_short_horizon_caps_ability(out_dir: Path) -> None:
    c = base_client()
    c["flujos"]["fases"] = [{"nombre": "Retiro", "anios": 2, "flujo_anual_real": -100_000}]
    res = write(out_dir, c)
    horizon = next(f for f in res["capacidad"]["factores"] if f["id"] == "horizonte")
    assert horizon["nivel"] == "bajo"
    assert calc.RANK[res["capacidad"]["nivel"]] <= calc.RANK["medio"]


def test_tolerance_is_min_of_ability_and_willingness(out_dir: Path) -> None:
    c = base_client()
    c["disposicion"] = {"likert": [{"id": "I1", "valor": 5, "reverse": False}],
                        "perdida_maxima_tolerada": 0.35, "comportamiento_pasado": "compro"}
    c["liquidez"]["necesidades_unicas"] = [{"descripcion": "Casa", "monto": 400_000, "meses": 12}]
    res = write(out_dir, c)
    assert res["disposicion"]["nivel"] == "alto"
    assert res["tolerancia"]["propuesta"] == calc.min_level(res["capacidad"]["nivel"], "alto")
    assert res["tolerancia"]["limite_perdida_pct"] <= 0.35


def test_willingness_contradiction_flagged(out_dir: Path) -> None:
    c = base_client()
    c["disposicion"] = {"likert": [{"id": "I1", "valor": 5, "reverse": False}],
                        "perdida_maxima_tolerada": 0.30, "comportamiento_pasado": "vendio"}
    res = write(out_dir, c)
    assert res["disposicion"]["contradiccion"] is True


def test_likert_out_of_range_rejected(out_dir: Path) -> None:
    c = base_client()
    c["disposicion"]["likert"] = [{"id": "I1", "valor": 7, "reverse": False}]
    (out_dir / "client.json").write_text(json.dumps(c), encoding="utf-8")
    with pytest.raises(calc.InputError):
        calc.compute(out_dir)


# ---------------------------------------------------------------- SAA check

def write_single_class_saa(out: Path, expected: float) -> None:
    cma = {"fecha": "2026-10-31", "fuente": "test", "base": "nominal",
           "clases": [{"clase": "A", "retorno_esperado": expected, "volatilidad": 0.01}],
           "correlaciones": [[1.0]]}
    saa = {"clases": [{"clase": "A", "objetivo": 1.0, "min": 1.0, "max": 1.0, "rol": "test"}]}
    (out / "cma.json").write_text(json.dumps(cma), encoding="utf-8")
    (out / "saa.json").write_text(json.dumps(saa), encoding="utf-8")


def test_saa_return_test_is_after_tax(out_dir: Path) -> None:
    # Regression: a mix that clears the required return before tax but not after
    # tax must fail the return test (eval run "empresaria-concentrada-cma").
    c = base_client()
    c["impuestos"]["tasa_sobre_rendimientos"] = 0.20
    c["supuestos_cliente"] = {"comision_anual": 0.01}
    required = write(out_dir, c)["retorno"]["nominal_pct"]
    write_single_class_saa(out_dir, required + 0.01 + 0.005)
    res = write(out_dir, c)["saa_check"]
    assert res["retorno_neto_pct"] >= required
    assert res["retorno_despues_impuestos_pct"] == pytest.approx(res["retorno_neto_pct"] * 0.8)
    assert res["cumple_retorno"] is False


def test_saa_return_test_undecided_without_tax(out_dir: Path) -> None:
    c = base_client()
    c["supuestos_cliente"] = {"comision_anual": 0.01}
    write_single_class_saa(out_dir, 0.30)
    res = write(out_dir, c)
    assert res["saa_check"]["cumple_retorno"] is None
    assert any("impuestos" in p and "Asignación" in p for p in res["pendientes"])


# ---------------------------------------------------------------- builder

def minimal_ips() -> dict[str, Any]:
    p = {"tipo": "parrafo", "texto": "Texto."}
    secs: list[dict[str, Any]] = [{"id": sid, "titulo": t, "bloques": [p], "usa": []}
                                  for sid, t in builder.REQUIRED_SECTIONS.items()]
    by_id = {s["id"]: s for s in secs}
    by_id["3"]["usa"] = ["perfil.edad", "perfil.edad_retiro", "perfil.fuente_riqueza", "situacion.ingreso_anual",
                         "situacion.estabilidad_ingreso", "situacion.patrimonio_neto", "situacion.deuda_total"]
    by_id["4"]["bloques"] = [{"tipo": "parrafo", "texto": "Real {{calc.retorno.real_pct}} sobre "
                              "{{client.situacion.cartera_invertible}}."}, {"tipo": "tabla_retorno"}]
    by_id["4"]["usa"] = ["flujos.fases", "flujos.meta_final_real"]
    by_id["5"]["bloques"] = [{"tipo": "tabla_capacidad"}, {"tipo": "tabla_disposicion"}]
    by_id["5"]["usa"] = ["disposicion.likert", "disposicion.perdida_maxima_tolerada",
                         "disposicion.comportamiento_pasado", "flexibilidad.puede_reducir_gasto",
                         "flexibilidad.puede_retrasar_meta", "flexibilidad.tiene_otros_ingresos"]
    by_id["6"]["bloques"] = [{"tipo": "subtitulo", "texto": k} for k in
                             ("Horizonte", "Impuestos", "Liquidez", "Legal y regulatorio", "Circunstancias únicas")]
    by_id["6"]["bloques"].append({"tipo": "tabla_liquidez"})
    by_id["6"]["usa"] = ["impuestos.residencia_fiscal", "liquidez.meses_reserva", "liquidez.reserva_actual",
                         "situacion.gasto_anual", "unicas"]
    by_id["7"]["bloques"] = [{"tipo": "parrafo", "texto": "[PENDIENTE: SAA del PM]"}]
    return {"secciones": secs, "no_usados": []}


def test_build_passes_and_writes_docx(out_dir: Path) -> None:
    write(out_dir, base_client())
    (out_dir / "ips.json").write_text(json.dumps(minimal_ips(), ensure_ascii=False), encoding="utf-8")
    assert builder.main(["x", str(out_dir)]) == 0
    assert (out_dir / "IPS_cliente-prueba_2026-11-19.docx").exists()


def test_build_fails_on_unused_fact(out_dir: Path) -> None:
    write(out_dir, base_client())
    ips = minimal_ips()
    ips["secciones"][2]["usa"].remove("situacion.deuda_total")
    (out_dir / "ips.json").write_text(json.dumps(ips, ensure_ascii=False), encoding="utf-8")
    assert builder.main(["x", str(out_dir)]) == 1


def test_build_fails_on_null_placeholder(out_dir: Path) -> None:
    write(out_dir, base_client())
    ips = minimal_ips()
    ips["secciones"][3]["bloques"][0]["texto"] = "Antes de impuestos {{calc.retorno.antes_impuestos_pct}}"
    (out_dir / "ips.json").write_text(json.dumps(ips, ensure_ascii=False), encoding="utf-8")
    assert builder.main(["x", str(out_dir)]) == 1


def test_final_requires_no_pending(out_dir: Path) -> None:
    write(out_dir, base_client())
    (out_dir / "ips.json").write_text(json.dumps(minimal_ips(), ensure_ascii=False), encoding="utf-8")
    assert builder.main(["x", str(out_dir), "--final"]) == 1
