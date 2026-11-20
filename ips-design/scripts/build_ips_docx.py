"""Validate ips.json against client.json / calc.json and render the IPS as .docx.

    python build_ips_docx.py <out_dir> [--final]

Writes <out_dir>/IPS_<slug>_<fecha_ips>.docx. Exit 1 (and no docx) when validation
fails: missing sections or auto-tables, unresolved placeholders, invalid `usa` paths,
facts neither used nor justified in `no_usados`. --final removes the draft header and
requires zero open items. Console output is ASCII only.
"""
from __future__ import annotations

import json
import re
import sys
import unicodedata
from collections.abc import Iterator, Sequence
from pathlib import Path
from typing import Any

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor
from docx.table import Table
from docx.text.paragraph import Paragraph

REQUIRED_SECTIONS: dict[str, str] = {
    "1": "Identificación",
    "2": "Propósito y gobierno",
    "3": "Perfil del inversionista",
    "4": "Objetivo de retorno",
    "5": "Objetivo de riesgo",
    "6": "Restricciones",
    "7": "Asignación estratégica",
    "8": "Lineamientos de inversión",
    "9": "Monitoreo, rebalanceo y revisión",
}
REQUIRED_TABLES: dict[str, tuple[str, ...]] = {
    "4": ("tabla_retorno",),
    "5": ("tabla_capacidad", "tabla_disposicion"),
    "6": ("tabla_liquidez",),
}
CONSTRAINT_KEYWORDS: tuple[str, ...] = ("horizonte", "impuestos", "liquidez", "legal", "circunstancias")
EXCLUDED_ROOTS: frozenset[str] = frozenset({"meta", "fuentes", "supuestos_cliente"})
DRAFT_MARK = "BORRADOR — sujeto a revisión de cumplimiento"
PLACEHOLDER = re.compile(r"\{\{(calc|client)\.([^}]+)\}\}")
INLINE = re.compile(r"(\*\*.+?\*\*|\[PENDIENTE:[^\]]*\])")
PENDING = re.compile(r"\[PENDIENTE:\s*([^\]]*)\]")
TABLE_STYLE = "Light Grid Accent 1"


class BuildError(ValueError):
    """Validation failure: the docx is not written."""


# ---------------------------------------------------------------- utilities

def fold(text: str) -> str:
    """Lowercase, accent-free text for keyword matching."""
    norm = unicodedata.normalize("NFKD", text)
    return "".join(c for c in norm if not unicodedata.combining(c)).lower()


def walk(obj: Any, dotted: str) -> tuple[bool, Any]:
    cur = obj
    for key in dotted.split("."):
        if isinstance(cur, dict) and key in cur:
            cur = cur[key]
        elif isinstance(cur, list) and key.isdigit() and int(key) < len(cur):
            cur = cur[int(key)]
        else:
            return False, None
    return True, cur


def fmt_value(root: str, path: str, value: Any) -> str:
    if isinstance(value, bool):
        return "sí" if value else "no"
    if isinstance(value, (int, float)):
        if root == "calc" and path.endswith("_pct"):
            return f"{value * 100:.2f}%"
        if isinstance(value, float) and 0 < abs(value) < 1:
            return f"{value * 100:.2f}%"
        if isinstance(value, int) and abs(value) < 10000:
            return str(value)
        return f"{value:,.0f}"
    return str(value)


def client_leaves(client: dict[str, Any]) -> dict[str, Any]:
    """Dotted path -> value for every non-null fact (lists are leaves)."""
    leaves: dict[str, Any] = {}

    def rec(obj: Any, prefix: str) -> None:
        if isinstance(obj, dict):
            for k, v in obj.items():
                rec(v, f"{prefix}.{k}" if prefix else k)
        elif obj is None or obj == [] or obj == "":
            return
        else:
            leaves[prefix] = obj

    for root, val in client.items():
        if root not in EXCLUDED_ROOTS:
            rec(val, root)
    return leaves


def covers(ref: str, leaf: str) -> bool:
    return leaf == ref or (ref.count(".") >= 1 and leaf.startswith(ref + "."))


def iter_texts(section: dict[str, Any]) -> Iterator[str]:
    for b in section.get("bloques", []):
        if "texto" in b:
            yield str(b["texto"])
        for item in b.get("items", []):
            yield str(item)
        for row in b.get("filas", []):
            for cell in row:
                yield str(cell)
        for h in b.get("encabezados", []):
            yield str(h)


# ---------------------------------------------------------------- validation

def resolve(text: str, data: dict[str, Any], errors: list[str], where: str) -> str:
    def sub(m: re.Match[str]) -> str:
        root, path = m.group(1), m.group(2)
        found, value = walk(data[root], path)
        if not found:
            errors.append(f"{where}: placeholder {{{{{root}.{path}}}}} does not exist")
            return m.group(0)
        if value is None:
            errors.append(f"{where}: placeholder {{{{{root}.{path}}}}} is null; write [PENDIENTE: ...] instead")
            return m.group(0)
        return fmt_value(root, path, value)

    return PLACEHOLDER.sub(sub, text)


def validate(client: dict[str, Any], calc: dict[str, Any], ips: dict[str, Any]) -> tuple[list[str], dict[str, list[str]]]:
    errors: list[str] = []
    sections = {str(s.get("id")): s for s in ips.get("secciones", [])}
    for sid, title in REQUIRED_SECTIONS.items():
        if sid not in sections:
            errors.append(f"section {sid} ({title}) missing")
    for sid, tables in REQUIRED_TABLES.items():
        types = {b.get("tipo") for b in sections.get(sid, {}).get("bloques", [])}
        for t in tables:
            if t not in types:
                errors.append(f"section {sid} must include block '{t}'")
    if "6" in sections:
        subs = [fold(str(b.get("texto", ""))) for b in sections["6"].get("bloques", []) if b.get("tipo") == "subtitulo"]
        for kw in CONSTRAINT_KEYWORDS:
            if not any(kw in s for s in subs):
                errors.append(f"section 6 needs a 'subtitulo' for '{kw}'")
    if "7" in sections:
        types7 = {b.get("tipo") for b in sections["7"].get("bloques", [])}
        text7 = " ".join(iter_texts(sections["7"]))
        if calc.get("saa_check") and "tabla_saa" not in types7:
            errors.append("section 7 must include 'tabla_saa' (calc.saa_check exists)")
        if not calc.get("saa_check") and "[PENDIENTE" not in text7:
            errors.append("section 7 has no CMA/SAA: it must carry a [PENDIENTE: ...] for the PM")

    data = {"calc": calc, "client": client}
    leaves = client_leaves(client)
    usage: dict[str, list[str]] = {leaf: [] for leaf in leaves}
    for sid, s in sections.items():
        refs: list[str] = list(s.get("usa", []))
        for text in iter_texts(s):
            resolve(text, data, errors, f"section {sid}")
            refs.extend(m.group(2) for m in PLACEHOLDER.finditer(text) if m.group(1) == "client")
        for ref in s.get("usa", []):
            if not walk(client, ref)[0]:
                errors.append(f"section {sid}: usa path '{ref}' not in client.json")
        for leaf in leaves:
            if any(covers(r, leaf) for r in refs) and sid not in usage[leaf]:
                usage[leaf].append(sid)
    justified = [str(n.get("path")) for n in ips.get("no_usados", []) if n.get("razon")]
    for leaf, used_in in usage.items():
        if not used_in and not any(covers(j, leaf) for j in justified):
            errors.append(f"fact '{leaf}' is not used in any section nor justified in no_usados")
    return errors, usage


# ---------------------------------------------------------------- rendering

def add_rich(par: Paragraph, text: str) -> None:
    for part in INLINE.split(text):
        if not part:
            continue
        if part.startswith("**") and part.endswith("**"):
            par.add_run(part[2:-2]).bold = True
        elif part.startswith("[PENDIENTE"):
            run = par.add_run(part)
            run.bold = True
            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
        else:
            par.add_run(part)


def add_table(doc: Any, headers: Sequence[str], rows: Sequence[Sequence[str]]) -> Table:
    table = doc.add_table(rows=1, cols=len(headers))
    try:
        table.style = TABLE_STYLE
    except (KeyError, ValueError):
        table.style = "Table Grid"
    for cell, h in zip(table.rows[0].cells, headers, strict=True):
        cell.text = ""
        cell.paragraphs[0].add_run(str(h)).bold = True
    for row in rows:
        cells = table.add_row().cells
        for cell, val in zip(cells, row, strict=False):  # short rows leave trailing cells empty
            cell.text = ""
            add_rich(cell.paragraphs[0], str(val))
    doc.add_paragraph()
    return table


def level(x: Any) -> str:
    return "n.d." if x is None else str(x).capitalize()


def table_retorno(doc: Any, calc: dict[str, Any]) -> None:
    r, d, sup = calc.get("retorno"), calc.get("retorno_deseado"), calc["supuestos"]
    if r is None:
        add_table(doc, ["Concepto", "Valor"], [["Retorno requerido", "[PENDIENTE: datos insuficientes]"]])
        return
    cols = ["Concepto", "Requerido"] + (["Deseado"] if d else [])

    def row(label: str, key: str) -> list[str]:
        return [label, str(r[key])] + ([str(d[key])] if d else [])

    rows = [
        row("Cartera inicial", "cartera_inicial_fmt"),
        row("Horizonte (años)", "horizonte_anios"),
        row("Meta final (dinero de hoy)", "meta_final_fmt"),
        row("Retorno real", "real_fmt"),
        ["Inflación supuesta", sup["inflacion"]["valor_fmt"]] + ([sup["inflacion"]["valor_fmt"]] if d else []),
        row("Retorno nominal (neto de comisiones)", "nominal_fmt"),
        ["Comisión anual", sup["comision_anual"]["valor_fmt"]] + ([sup["comision_anual"]["valor_fmt"]] if d else []),
        row("Retorno nominal bruto de comisiones", "bruto_comisiones_fmt"),
        ["Tasa de impuestos sobre rendimientos", sup["tasa_sobre_rendimientos"]["valor_fmt"]]
        + ([sup["tasa_sobre_rendimientos"]["valor_fmt"]] if d else []),
        row("Retorno antes de impuestos", "antes_impuestos_fmt"),
    ]
    add_table(doc, cols, rows)
    if calc.get("trayectoria"):
        add_table(doc, ["Fase", "Años", "Flujo anual (dinero de hoy)", "Riqueza al inicio"],
                  [[p["fase"], str(p["anios"]), p["flujo_anual_real_fmt"], p["riqueza_inicio_fmt"]]
                   for p in calc["trayectoria"]])
    p = doc.add_paragraph()
    p.add_run("Método: ").italic = True
    p.add_run(r["metodo"]).italic = True


def table_capacidad(doc: Any, calc: dict[str, Any]) -> None:
    c = calc["capacidad"]
    rows = [[f["etiqueta"], f["valor_fmt"], level(f["nivel"]), "Sí" if f["core"] else "No"] for f in c["factores"]]
    rows.append(["**Mediana de factores**", "", level(c["mediana"]), ""])
    rows.append(["**Tope por factores clave en nivel bajo**", "", level(c["tope"]) if c["tope"] else "Sin tope", ""])
    rows.append(["**Capacidad propuesta**", "", f"**{level(c['nivel'])}**", ""])
    add_table(doc, ["Factor de capacidad", "Dato", "Nivel", "Clave"], rows)


def table_disposicion(doc: Any, calc: dict[str, Any]) -> None:
    w, t, c = calc["disposicion"], calc["tolerancia"], calc["capacidad"]
    rows = [[s["etiqueta"], s["valor_fmt"], level(s["nivel"])] for s in w["fuentes"]]
    rows.append(["**Disposición propuesta**", "Contradicción entre fuentes: " + ("sí" if w["contradiccion"] else "no"),
                 f"**{level(w['nivel'])}**"])
    add_table(doc, ["Fuente de disposición", "Dato", "Nivel"], rows)
    add_table(doc, ["Conclusión", "Valor"], [
        ["Capacidad", level(c["nivel"])],
        ["Disposición", level(w["nivel"])],
        ["**Tolerancia propuesta = mín(capacidad, disposición)**", f"**{level(t['propuesta'])}**"],
        ["Pérdida máxima en un año (límite)", t["limite_perdida_fmt"]],
        ["Decisión del PM", "________________________"],
    ])


def table_liquidez(doc: Any, calc: dict[str, Any]) -> None:
    q = calc["liquidez"]
    add_table(doc, ["Concepto", "Valor"], [
        ["Reserva de emergencia requerida", q["reserva_requerida_fmt"]],
        ["Reserva disponible hoy", q["reserva_actual_fmt"]],
        ["Faltante de reserva", q["faltante_reserva_fmt"]],
        [f"Necesidades extraordinarias en {q['horizonte_meses']} meses", q["necesidades_horizonte_fmt"]],
        ["(Faltante + necesidades) / cartera", q["ratio_fmt"]],
    ])


def table_saa(doc: Any, calc: dict[str, Any]) -> None:
    s = calc["saa_check"]
    add_table(doc, ["Clase de activo", "Objetivo", "Rango permitido", "Rol"],
              [[c["clase"], c["objetivo_fmt"], c["rango_fmt"], str(c.get("rol", ""))] for c in s["clases"]])

    def yes_no(x: Any) -> str:
        return "n.d." if x is None else ("Sí" if x else "No")

    add_table(doc, ["Métrica (CMA " + str(s.get("cma_fecha")) + ")", "Valor"], [
        ["Retorno esperado (bruto)", s["retorno_esperado_fmt"]],
        ["Retorno esperado neto de comisiones", s["retorno_neto_fmt"]],
        ["Retorno esperado neto de comisiones e impuestos", s.get("retorno_despues_impuestos_fmt", "n.d.")],
        ["Retorno nominal requerido (después de comisiones e impuestos)", s.get("retorno_requerido_nominal_fmt", "n.d.")],
        ["Volatilidad esperada", s["volatilidad_fmt"]],
        ["Shortfall (E[R] neto − z·σ)", s["shortfall_fmt"]],
        ["¿Alcanza el retorno requerido (después de impuestos)?", yes_no(s["cumple_retorno"])],
        ["¿Respeta el límite de pérdida?", yes_no(s["cumple_riesgo"])],
    ])


AUTO_TABLES = {
    "tabla_retorno": table_retorno,
    "tabla_capacidad": table_capacidad,
    "tabla_disposicion": table_disposicion,
    "tabla_liquidez": table_liquidez,
    "tabla_saa": table_saa,
}


def add_page_field(par: Paragraph, instr: str) -> None:
    run = par.add_run()
    for tag, text in (("begin", None), (None, instr), ("end", None)):
        if tag:
            el = OxmlElement("w:fldChar")
            el.set(qn("w:fldCharType"), tag)
        else:
            el = OxmlElement("w:instrText")
            el.set(qn("xml:space"), "preserve")
            el.text = text
        run._r.append(el)


def setup(doc: Any, client: dict[str, Any], final: bool) -> None:
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(10.5)
    sec = doc.sections[0]
    head = sec.header.paragraphs[0]
    head.text = ""
    if not final:
        run = head.add_run(DRAFT_MARK)
        run.bold = True
        run.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
        head.add_run("   ")
    head.add_run(f"IPS · {client['meta']['cliente']}").font.size = Pt(9)
    foot = sec.footer.paragraphs[0]
    foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
    foot.add_run("Página ")
    add_page_field(foot, "PAGE")
    foot.add_run(" de ")
    add_page_field(foot, "NUMPAGES")


def render(client: dict[str, Any], calc: dict[str, Any], ips: dict[str, Any],
           usage: dict[str, list[str]], final: bool, out_file: Path) -> list[str]:
    meta = client["meta"]
    data = {"calc": calc, "client": client}
    sink: list[str] = []
    doc = Document()
    setup(doc, client, final)

    doc.add_heading("Declaración de Política de Inversión", 0)
    add_table(doc, ["Campo", "Valor"], [
        ["Cliente", meta["cliente"]],
        ["Tipo de inversionista", str(meta.get("tipo_inversionista", "individual")).capitalize()],
        ["Asesor", str(meta.get("asesor") or "[PENDIENTE: asesor]")],
        ["Fecha del IPS", str(meta["fecha_ips"])],
        ["Fecha de corte de los datos", str(meta.get("fecha_datos") or "[PENDIENTE: fecha de corte]")],
        ["Moneda", str(meta.get("moneda", ""))],
        ["Versión", "Final" if final else "Borrador"],
    ])

    sections = sorted(ips["secciones"], key=lambda s: int(str(s["id"])) if str(s["id"]).isdigit() else 99)
    for s in sections:
        sid = str(s["id"])
        doc.add_heading(f"{sid}. {REQUIRED_SECTIONS.get(sid, s.get('titulo', ''))}", 1)
        for b in s.get("bloques", []):
            kind = b.get("tipo")
            if kind in AUTO_TABLES:
                AUTO_TABLES[kind](doc, calc)
            elif kind == "parrafo":
                add_rich(doc.add_paragraph(), resolve(b["texto"], data, sink, sid))
            elif kind == "subtitulo":
                doc.add_heading(resolve(b["texto"], data, sink, sid), 2)
            elif kind == "vinetas":
                for item in b.get("items", []):
                    add_rich(doc.add_paragraph(style="List Bullet"), resolve(item, data, sink, sid))
            elif kind == "tabla":
                add_table(doc, [resolve(h, data, sink, sid) for h in b["encabezados"]],
                          [[resolve(str(c), data, sink, sid) for c in row] for row in b["filas"]])
            else:
                raise BuildError(f"section {sid}: unknown block type '{kind}'")

    doc.add_heading("10. Firmas", 1)
    doc.add_paragraph(
        "Las partes declaran que la información contenida en esta política es correcta a la fecha de "
        "corte y que la cartera se administrará conforme a ella hasta su siguiente revisión.")
    add_table(doc, ["Parte", "Nombre", "Firma", "Fecha"], [
        ["Cliente", meta["cliente"], "", ""],
        ["Asesor", str(meta.get("asesor") or ""), "", ""],
    ])

    doc.add_heading("Anexos", 1)
    doc.add_heading("A. Supuestos", 2)
    sup_rows = [[k.replace("_", " ").capitalize(), v["valor_fmt"], v["fuente"]] for k, v in calc["supuestos"].items()]
    for note in (calc.get("retorno") or {}).get("notas", []):
        sup_rows.append(["Retorno requerido", "", note])
    sup_rows += [[x["supuesto"], x["valor"], x["fuente"]] for x in ips.get("supuestos", [])]
    add_table(doc, ["Supuesto", "Valor", "Fuente"], sup_rows)

    doc.add_heading("B. Pendientes", 2)
    pend = collect_pending(calc, ips)
    add_table(doc, ["#", "Pendiente"], [[str(i + 1), p] for i, p in enumerate(pend)] or [["—", "Ninguno"]])

    doc.add_heading("C. Trazabilidad de la información del cliente", 2)
    fuentes = client.get("fuentes", {})
    justified = {str(n["path"]): str(n["razon"]) for n in ips.get("no_usados", [])}
    rows = []
    for leaf, used_in in usage.items():
        where = ", ".join(sorted(used_in, key=int)) if used_in else "No usado: " + next(
            (r for p, r in justified.items() if covers(p, leaf)), "")
        rows.append([leaf, str(fuentes.get(leaf, "")), where])
    add_table(doc, ["Dato", "Evidencia (respuesta del cliente)", "Secciones"], rows)

    if ips.get("glosario"):
        doc.add_heading("D. Glosario", 2)
        add_table(doc, ["Término", "Definición"], [[g["termino"], g["definicion"]] for g in ips["glosario"]])

    if sink:
        raise BuildError("; ".join(sink))
    doc.save(str(out_file))
    return pend


def collect_pending(calc: dict[str, Any], ips: dict[str, Any]) -> list[str]:
    items: list[str] = list(calc.get("pendientes", []))
    items += [f"Sección {p.get('seccion')}: {p['texto']}" for p in ips.get("pendientes", [])]
    for s in ips.get("secciones", []):
        for text in iter_texts(s):
            items += [f"Sección {s['id']}: {m.group(1).strip()}" for m in PENDING.finditer(text)]
    seen: set[str] = set()
    unique = []
    for it in items:
        key = fold(it)
        if key not in seen:
            seen.add(key)
            unique.append(it)
    return unique


def main(argv: Sequence[str]) -> int:
    args = [a for a in argv[1:] if not a.startswith("--")]
    final = "--final" in argv
    if len(args) != 1:
        print("usage: python build_ips_docx.py <out_dir> [--final]")
        return 2
    out = Path(args[0])
    try:
        client = json.loads((out / "client.json").read_text(encoding="utf-8"))
        calc = json.loads((out / "calc.json").read_text(encoding="utf-8"))
        ips = json.loads((out / "ips.json").read_text(encoding="utf-8"))
    except (FileNotFoundError, json.JSONDecodeError) as exc:
        print(f"[x] FAIL: cannot read inputs: {exc}".encode("ascii", "replace").decode())
        return 1
    errors, usage = validate(client, calc, ips)
    if final and collect_pending(calc, ips):
        errors.append("--final requires zero open items (PENDIENTE)")
    if errors:
        print(f"[x] FAIL: {len(errors)} validation error(s)")
        for e in errors:
            print("    - " + e.encode("ascii", "replace").decode())
        return 1
    slug = client["meta"].get("slug") or "cliente"
    out_file = out / f"IPS_{slug}_{client['meta']['fecha_ips']}.docx"
    try:
        pend = render(client, calc, ips, usage, final, out_file)
    except BuildError as exc:
        print("[x] FAIL: " + str(exc).encode("ascii", "replace").decode())
        return 1
    used = sum(1 for v in usage.values() if v)
    print(f"[ok] PASS: wrote {out_file.name}".encode("ascii", "replace").decode())
    print(f"     facts used: {used}/{len(usage)}  open items: {len(pend)}  draft: {not final}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
