<div align="center">

# Skills

**Repositorio oficial de agent skills de CFA Society Mexico**
*Skills abiertas para análisis de inversiones, construidas sobre el currículum CFA y mantenidas por voluntarios.*

[![License: MIT](https://img.shields.io/badge/licencia-MIT-blue.svg)](LICENSE)
[![Skills](https://img.shields.io/badge/skills-2-0A66C2.svg)](#catálogo)
[![Agent Skills](https://img.shields.io/badge/formato-Agent%20Skills-8A2BE2.svg)](#instalación)
[![Claude Code](https://img.shields.io/badge/Claude%20Code-compatible-D97757.svg)](#instalación)
[![Python](https://img.shields.io/badge/python-3.10%2B-3776AB.svg?logo=python&logoColor=white)](#instalación)
[![PRs bienvenidos](https://img.shields.io/badge/PRs-bienvenidos-brightgreen.svg)](#contribuir)
[![Último commit](https://img.shields.io/github/last-commit/CFA-Society-Mexico/Skills.svg)](https://github.com/CFA-Society-Mexico/Skills/commits/main)

[Catálogo](#catálogo) · [Instalación](#instalación) · [Principios](#principios-de-diseño) · [Contribuir](#contribuir) · [Licencia](#licencia-y-avisos)

</div>

---

## Catálogo

Cada carpeta es una skill autocontenida (`SKILL.md` + referencias + scripts) en el formato
[Agent Skills](https://platform.claude.com/docs/en/agents-and-tools/agent-skills/overview), que
pueden cargar Claude Code y otros agentes compatibles.

| Skill | Qué hace | Entrega | Marco |
|---|---|---|---|
| 📊 [**credit-4c**](credit-4c/) | Análisis de crédito de una empresa a partir de sus estados financieros, con ajustes de deuda estilo S&P, estrés y revisión de liquidez y vencimientos | Excel con fórmulas vivas + reporte PDF con gráficas | CFA 4C: Capacidad, Colateral, Covenants, Carácter |
| 📝 [**ips-design**](ips-design/) | Redacta la Declaración de Política de Inversión (IPS) de un cliente individual a partir de sus respuestas al cuestionario de perfil | Documento Word (.docx) en español, en borrador para firma | CFA Institute: objetivos y restricciones (RRTTLLU), capacidad vs disposición al riesgo |

> [!TIP]
> **¿Usas agentes en tu trabajo de inversión? Este catálogo crece con la comunidad.**
> Si tienes un proceso que repites (una valuación, un reporte, un análisis), conviértelo en
> skill y compártelo. [Así puedes colaborar →](#contribuir)

---

## Instalación

Clona el repositorio y copia la carpeta de la skill que quieras en tu directorio de skills.

**Windows (PowerShell)**
```powershell
git clone https://github.com/CFA-Society-Mexico/Skills.git
Copy-Item -Recurse Skills\ips-design "$HOME\.claude\skills\ips-design"
```

**macOS / Linux**
```bash
git clone https://github.com/CFA-Society-Mexico/Skills.git
cp -r Skills/ips-design ~/.claude/skills/ips-design
```

Después pídele al agente la tarea en lenguaje natural; la skill se activa sola.

| Skill | Requisitos |
|---|---|
| credit-4c | Python 3.10+ con `openpyxl`, `pdfplumber`, `matplotlib`, `reportlab`, `pyyaml`; Excel (Windows, vía `pywin32`) o LibreOffice para recalcular |
| ips-design | Python 3.10+ con `python-docx` (y `pytest` para correr las pruebas) |

---

## Principios de diseño

Todas las skills del repositorio están construidas contra las fallas conocidas de los
analistas y asesores basados en LLM:

| Principio | En la práctica |
|---|---|
| 🧑‍⚖️ **El modelo propone, el profesional decide** | Calificaciones, tolerancias y asignaciones son *propuestas*; las celdas o líneas de decisión quedan vacías para el analista o PM |
| 🧮 **El LLM no hace aritmética** | Los números los calculan fórmulas de Excel o scripts de Python; el texto los cita, no los reescribe |
| 🔎 **Nada sin fuente** | Cada cifra capturada lleva documento y página, o la cita textual del cliente; lo desconocido queda como `[PENDIENTE]`, nunca como un supuesto silencioso |
| 🚦 **Puertas de aprobación** | El flujo se detiene para que apruebes la captura de datos antes de construir el entregable |
| ⚖️ **Contra el colapso heurístico** | Ninguna métrica aislada decide el resultado: los factores se evalúan por separado y se muestran todos |

---

## Skills

### 📊 credit-4c

Análisis de crédito de una empresa con el marco 4C del CFA (Capacidad, Colateral,
Covenants, Carácter), ajustes de deuda estilo S&P, prueba de estrés y revisión de liquidez
y vencimientos.

**Entrega**
- Libro de Excel en el que cada cifra fuera de la tabla citada `Sources` es una fórmula viva
  (insumos en azul, cálculos en negro, vínculos en verde), con chequeos de integridad.
- Reporte PDF con gráficas, hallazgos por C, caso bajista y límites del análisis.

**Uso**: *"Analiza el crédito de &lt;emisora&gt; con el framework 4C; los estados financieros
están en &lt;carpeta&gt;"*. Las bandas de política por omisión
(`credit-4c/assets/policy-default.yaml`) vienen de las tablas de métricas de crédito de
CFA Nivel I más defaults documentados; reemplázalas con tu política antes de confiar en
cualquier propuesta.

**No es para**: bancos y aseguradoras, valuación de capital ni modelos de tres estados
completos.

### 📝 ips-design

Redacta la Declaración de Política de Inversión (IPS) de un cliente individual siguiendo
la práctica del CFA Institute. Cubre objetivos de retorno y riesgo; capacidad y disposición
al riesgo; restricciones de horizonte, impuestos, liquidez, legales y circunstancias únicas;
asignación estratégica y detonantes de revisión.

**Entrega**
- IPS en Word, en español, con marca de borrador en cada página, bloque de firmas y anexos:
  supuestos con fuente, pendientes y una tabla de trazabilidad que liga cada dato del
  cliente con las secciones que lo usan.
- Los archivos de trabajo: `client.json` (datos + citas), `calc.json` (cálculos) e
  `ips.json` (narrativa).

**Cómo funciona**
- La **capacidad** se evalúa factor por factor con datos duros: horizonte, tasa de retiro,
  liquidez, retorno requerido, estabilidad del ingreso, deuda y flexibilidad.
- La **disposición** se evalúa con tres fuentes independientes (cuestionario, pérdida
  tolerada y comportamiento pasado) y se señalan las contradicciones.
- La **tolerancia propuesta** es la menor de las dos, con una línea vacía para la decisión
  del PM.
- Con las expectativas de mercado (CMA) del PM propone una **asignación estratégica** y
  verifica que alcance el retorno requerido después de comisiones e impuestos.

**Uso**
- Sin respuestas todavía: pide el cuestionario en blanco (`ips-design/assets/questionnaire.md`).
- Con respuestas: *"Arma la IPS de &lt;cliente&gt;; sus respuestas están en &lt;archivo&gt;"*.
  Agrega tu CMA (formato `assets/cma-template.json`) para obtener una asignación propuesta.
- `build_ips_docx.py <carpeta> --final` quita la marca de borrador y solo corre con cero
  pendientes.

Las bandas de `assets/policy-default.json` son reglas de la casa para calibrar, no cifras
del CFA. Los requisitos regulatorios locales (por ejemplo, el perfil de inversión que
exige el regulador) quedan como lista de verificación para cumplimiento.

<details>
<summary><b>Limitaciones conocidas (v1)</b></summary>

- Una sola tolerancia para toda la cartera. No hay canastas por meta, así que una meta grande
  de corto plazo (por ejemplo, el enganche de una casa) también limita el dinero de largo plazo.
- Un ingreso grande e incierto, como la venta de una empresa, no se puede modelar como
  escenario; se deja fuera de los flujos y se señala.
- El retorno deseado solo varía la meta al final del horizonte, no metas intermedias.
- Las líneas de firma se generan solo para cliente y asesor; los co-decisores se agregan en
  el texto.

</details>

**No es para**: inversionistas institucionales (fondos de pensiones, fundaciones,
endowments, aseguradoras), selección de valores ni aprobación de cumplimiento.

---

## Contribuir

### 🤝 Te invitamos a sumar tus skills

Este repositorio es un proyecto abierto de la comunidad de CFA Society Mexico. Si eres
charterholder, candidato o profesional de inversiones y tienes un proceso que haces una y
otra vez, probablemente ya es una skill esperando ser escrita. Compártela: tu trabajo puede
ahorrarle horas a cientos de colegas y elevar el estándar de cómo la industria usa la IA.

**Ideas que nos encantaría ver**

| Área | Ejemplos de skills |
|---|---|
| Renta variable | Valuación DCF con reverse DCF, múltiplos comparables, revisión de resultados trimestrales |
| Renta fija | Análisis de curva, duración y convexidad de cartera, valor relativo entre emisiones |
| Portafolios | Rebalanceo por corredores, atribución de desempeño, presentación conforme a GIPS® |
| Riesgo | VaR y pruebas de estrés, análisis de concentración, monitoreo de límites |
| Ética y cumplimiento | Revisión de idoneidad, checklist del Código y Estándares, conflictos de interés |
| Mercado mexicano | Lectura de reportes BMV, análisis de FIBRAs, curvas de Cetes y Mbonos |

¿Tienes otra idea? Propónla en un
[issue](https://github.com/CFA-Society-Mexico/Skills/issues) antes de escribirla y
la afinamos juntos.

### Cómo colaborar

1. Haz un fork del repositorio y crea una rama para tu skill.
2. Pon tu skill en su propia carpeta, con `SKILL.md` más `references/`, `scripts/` y
   `assets/` según necesite.
3. Pon los cálculos en scripts con pruebas (`pytest`). Los scripts imprimen PASS/FAIL y salen
   con código distinto de cero si algo falla.
4. Aplica los [principios de diseño](#principios-de-diseño): el modelo propone, el
   profesional decide, y ninguna cifra queda sin fuente.
5. No subas datos de clientes reales, rutas locales ni credenciales.
6. Agrega tu skill al [catálogo](#catálogo) y abre un pull request.

¿Encontraste un error? Abre un
[issue](https://github.com/CFA-Society-Mexico/Skills/issues).

---

## Licencia y avisos

Distribuido bajo la [licencia MIT](LICENSE) · © 2026 CFA Society Mexico.

> [!IMPORTANT]
> **No es asesoría de inversión.** Estas skills generan borradores y propuestas para
> profesionales de inversión; no sustituyen el juicio profesional, la revisión de
> cumplimiento ni la asesoría fiscal o legal. Verifica cada resultado antes de usarlo con un
> cliente.

CFA®, Chartered Financial Analyst® y GIPS® son marcas registradas propiedad de CFA Institute.
