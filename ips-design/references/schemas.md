# File schemas

Three JSON files live in the output folder: `client.json` (you write it from the answers),
`calc.json` (`ips_calc.py` writes it), and `ips.json` (you write the narrative). `policy.json`
is the copied `assets/policy-default.json`. `cma.json` and `saa.json` are optional.

Unknown values are `null`, never a guess. Amounts are plain numbers in the account currency
(no thousands separators); rates are decimals (`0.05` = 5%).

## client.json
```json
{
  "meta": {
    "cliente": "Nombre Apellido",
    "slug": "nombre-apellido",
    "fecha_ips": "2026-11-19",
    "fecha_datos": "2026-10-31",
    "moneda": "MXN",
    "asesor": "Nombre del asesor",
    "tipo_inversionista": "individual"
  },
  "perfil": {
    "edad": 31, "edad_retiro": 65, "estado_civil": "soltero", "dependientes": 0,
    "ocupacion": "texto", "fuente_riqueza": "activa | pasiva | ahorro_empleo",
    "experiencia_inversion": "texto", "otras_partes": "texto o null"
  },
  "situacion": {
    "cartera_invertible": 2500000, "patrimonio_neto": 4000000, "otros_activos": "texto",
    "deuda_total": 300000, "ingreso_anual": 900000,
    "estabilidad_ingreso": "alta | media | baja", "gasto_anual": 480000
  },
  "metas": [
    {"descripcion": "Retiro", "monto_real": 600000, "anio": 2060, "prioridad": "critica | deseable"}
  ],
  "flujos": {
    "fases": [
      {"nombre": "Acumulación", "anios": 34, "flujo_anual_real": 120000},
      {"nombre": "Retiro", "anios": 30, "flujo_anual_real": -600000}
    ],
    "meta_final_real": 0
  },
  "deseado": {"meta_final_real": 5000000, "descripcion": "Herencia deseable"},
  "liquidez": {
    "meses_reserva": 6, "reserva_actual": 150000,
    "necesidades_unicas": [{"descripcion": "Enganche", "monto": 400000, "meses": 18}],
    "activos_iliquidos": "texto o null"
  },
  "impuestos": {
    "residencia_fiscal": "México", "tasa_marginal": null,
    "tasa_sobre_rendimientos": null, "vehiculos": "texto o null"
  },
  "legal": {"estructura": "texto o null", "restricciones": "texto o null"},
  "unicas": ["Excluir tabaco", "Posición concentrada en ..."],
  "disposicion": {
    "likert": [
      {"id": "I1", "valor": 4, "reverse": false},
      {"id": "I2", "valor": 2, "reverse": true}
    ],
    "perdida_maxima_tolerada": 0.15,
    "comportamiento_pasado": "vendio | mantuvo | compro | null"
  },
  "flexibilidad": {
    "puede_reducir_gasto": true, "puede_retrasar_meta": null, "tiene_otros_ingresos": false
  },
  "preferencias": {"reportes": "texto", "participacion": "texto"},
  "supuestos_cliente": {"inflacion": null, "comision_anual": null},
  "fuentes": {"situacion.cartera_invertible": "B2: 'tengo 2.5 millones en ...'"}
}
```
- `flujos.fases`: consecutive years, end-of-year flows **from the portfolio's view** in
  today's money: positive = contribution, negative = withdrawal. Net the withdrawal of other
  income (pension, rent) before writing it: if spending is 600k and a pension pays 200k, the
  flow is −400k.
- `flujos.meta_final_real`: wealth required at the end of the last phase, in today's money.
  `null` if the client didn't say; the calc then assumes 0 and lists it as an assumption to confirm.
- `deseado` is optional: a second, aspirational target for the desired return.
- `disposicion.likert`: questionnaire items I1–I6; I2, I3 and I6 are `reverse: true`.
- `fuentes`: one short quote per fact, keyed by dotted path. It is the evidence trail.
- `supuestos_cliente` overrides `policy.json` assumptions when the client or PM gives a value.

## calc.json (written by `ips_calc.py` — never edit)
Top-level keys: `as_of`, `supuestos`, `retorno`, `retorno_deseado`, `liquidez`,
`capacidad`, `disposicion`, `tolerancia`, `trayectoria`, `saa_check`, `pendientes`.
Fields ending in `_pct` are decimals rendered as percentages; `_fmt` fields are ready-to-print
strings. Reference any scalar from the IPS text with `{{calc.<path>}}`.

## saa.json (optional, only with cma.json)
```json
{"clases": [{"clase": "Renta variable global", "objetivo": 0.40, "min": 0.30, "max": 0.50, "rol": "Crecimiento"}]}
```
Class names must match `cma.json`. Targets must sum to 1.

## ips.json
```json
{
  "secciones": [
    {
      "id": "4",
      "titulo": "Objetivo de retorno",
      "bloques": [
        {"tipo": "parrafo", "texto": "El retorno real requerido es {{calc.retorno.real_pct}} ..."},
        {"tipo": "tabla_retorno"},
        {"tipo": "subtitulo", "texto": "Retorno deseado"},
        {"tipo": "vinetas", "items": ["...", "..."]},
        {"tipo": "tabla", "encabezados": ["A", "B"], "filas": [["1", "2"]]}
      ],
      "usa": ["flujos.fases", "flujos.meta_final_real"]
    }
  ],
  "no_usados": [{"path": "preferencias.participacion", "razon": "Sin impacto en la política"}],
  "pendientes": [{"seccion": "6", "texto": "Tasa de impuestos sobre rendimientos"}],
  "supuestos": [{"supuesto": "Inflación de largo plazo", "valor": "3.00%", "fuente": "Banxico"}],
  "glosario": [{"termino": "IPS", "definicion": "..."}]
}
```
- Required sections: ids `1`–`9` with the titles in `cfai-ips-framework.md`. Section 10
  (signatures) and the annexes are generated by the builder.
- Auto-tables (numbers come from `calc.json`, never typed): `tabla_retorno` (section 4),
  `tabla_capacidad` and `tabla_disposicion` (section 5), `tabla_liquidez` (section 6),
  `tabla_saa` (section 7, only when `calc.saa_check` exists).
- Placeholders: `{{calc.path}}` and `{{client.path}}` (amounts get thousands separators).
  An unresolved placeholder fails the build.
- `**negritas**` is supported. Any `[PENDIENTE: ...]` in the text is highlighted.
- Section 6 must contain five `subtitulo` blocks: Horizonte, Impuestos, Liquidez, Legal y
  regulatorio, Circunstancias únicas.
- Traceability: every non-null fact in `client.json` must be used (in a section's `usa` or a
  `{{client...}}` placeholder) or listed in `no_usados` with a reason. Coverage below 100%
  of that set fails the build.
