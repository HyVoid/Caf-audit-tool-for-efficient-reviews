[ 🌐 عربي ](README.ar.md) | [ 🇩🇪 Deutsch ](README.de.md) | [ 🇫🇷 Français ](README.fr.md) | [ 🇳🇱 Nederlands ](README.nl.md) | [ 🇪🇸 Español ](README.sp.md) | [ 🇬🇧 English ](README.md)

# Auditoría de cafetería · Motor de diagnóstico 3D

[![Licencia](https://img.shields.io/badge/License-Apache_2.0-blue.svg)](https://opensource.org/licenses/Apache-2.0)
[![Plataforma](https://img.shields.io/badge/Platform-Excel%20%7C%20Web-navy)](https://github.com)
[![Tipo de herramienta](https://img.shields.io/badge/Tool_Type-Decision_Support-2251FF)](https://github.com)

**Un marco de auditoría estructurado que convierte la observación in situ de la cafetería en una puntuación diagnóstica ponderada, una estimación de la fuga de ingresos y una lista de acciones priorizada — en menos de 30 minutos por visita.**

---

[▶ Demostración en vivo](https://hyvoid.github.io/Caf-audit-tool-for-efficient-reviews/) · [⬇ Descargar libro de trabajo](cafe_audit_tool.xlsx) · 

---

 <img width="1536" height="1024" alt="ChatGPT Image Jul 14, 2026, 06_20_33 PM" src="https://github.com/user-attachments/assets/99030acd-62a9-4f24-b7dd-425c643c8a39" />


---

## ¿Qué decisión le ayuda a tomar esta herramienta?

1. **Qué dimensión operativa está costando más ingresos** — Experiencia del Cliente, Operaciones u Oportunidad de Ingresos — ponderada por el tipo de tienda?
2. **Si esta cafetería rinde por encima o por debajo del umbral mínimo de la industria** — y en qué medida, expresado como una puntuación compuesta en una escala verificada de 5 puntos?
3. **Cuál es la fuga mensual de ingresos** derivada de fallos observables, y si esa cifra cruza el umbral de intervención de $500?
4. **Qué métricas específicas requieren acción inmediata** y cuál es el primer paso correctivo para cada una?

---

## Acerca del creador

Tres restricciones operativas gobiernan cada decisión de diseño de este libro de trabajo.

**Soporte de decisión por encima del reemplazo de sistemas.** Esto no es una integración con TPV, una plataforma de análisis de fidelización ni una suite de gestión de operaciones. Es un instrumento de observación estructurado que produce un resultado listo para la toma de decisiones a partir de datos de campo recopilados en una tableta durante una sola visita.

**Razonamiento productizado por encima de la acumulación de funciones.** Cada fórmula sirve a una decisión. El motor de ponderación SUMPRODUCT existe porque los promedios sin ponderar tergiversan de forma sistemática el rendimiento de cafeterías con distintos modelos operativos. La lista de acciones basada en FILTER existe porque un panel de control sin priorización sigue requiriendo juicio humano para actuar.

**Ejecución de baja fricción por encima de la complejidad de implementación.** Sin macros. Sin dependencias de VBA. Sin conexiones de datos externas. El libro de trabajo se ejecuta en Excel para iOS, Android, iPadOS y Excel Online sin modificación.

---

## Para quién es

| Usuario | Situación típica | Decisión que respalda | Por qué fallan los métodos existentes |
|---|---|---|---|
| Consultor de cafeterías independiente | Auditar 4–12 locales al mes para el propietario de una cartera | Qué locales necesitan capital, cuáles necesitan proceso y cuáles necesitan personal | Las visitas basadas en la intuición producen resultados inconsistentes; los clientes no pueden comparar entre visitas |
| Gerente de operaciones de alimentación y bebidas | Comparar el rendimiento de las sucursales para la revisión trimestral | Dónde enfocar la atención del gerente regional | Los datos de ventas por sucursal no aíslan la causa raíz; dos cafeterías con ingresos idénticos pueden tener una salud operativa opuesta |
| Propietario / operador de cafetería | Autodiagnóstico antes de una decisión de renovación de contrato de arrendamiento o de reforma | Si el problema es el producto, el personal o el espacio | La proximidad del propietario crea puntos ciegos; no hay un punto de referencia externo para las puntuaciones |
| Proveedor de formación hostelera | Evaluación antes/después de la formación para cafeterías cliente | Si la formación produjo una mejora operativa medible | La comparación antes/después requiere un instrumento de medición coherente, que la mayoría de los operadores no tienen |

---

## Por qué la mayoría de los errores de auditoría de cafeterías no son errores de juicio

El auditor suele tener experiencia suficiente. El error se introduce antes — en la estructura de medición.

**Dónde se rompe el proceso:**

```
Unstructured observation
    → inconsistent metric coverage per visit
    → no rubric → scores reflect auditor mood, not café performance
    → averages across dimensions hide the worst-performing metric
    → action list ranked by salience, not by weighted impact
    → owner receives a report; no decision follows
```

**Con un instrumento estructurado:**

```
Rubric-anchored 5-point scoring (12 metrics × 3 shifts)
    → AVERAGE per metric across shifts → removes single-event bias
    → SUMPRODUCT × store-type weights → composite reflects operating model
    → IFS grade → single comparable output
    → FILTER(Red_Flag_Vector = TRUE) → action list auto-extracts
    → Lost customers × ASP × 30 → converts observation to financial stakes
```

**Implicación operativa:** Un auditor que puntúa "Cola y tiempo de espera" con un 3 por intuición puede puntuarlo con un 2 cuando se ve obligado a responder: *"Is the wait 4–6 minutes with no proactive ETA communication?"* La rúbrica elimina la brecha entre dos auditores que observan la misma cola.

El libro de trabajo no emite el juicio. Evita que el juicio se emita sin una definición.

---

## Tres trampas que atrapan incluso a los operadores experimentados

### Trampa 1 — El promedio sin ponderar tergiversa el modelo operativo

**Se tomó una decisión.** Una cafetería puntúa 3,8 en Experiencia del Cliente y 2,9 en Operaciones. El promedio sin ponderar de las tres dimensiones es 3,30. El operador concluye que el rendimiento es aceptable y pospone la intervención.

**El defecto inadvertido.** La cafetería es una operación Speed-First centrada en el servicio para llevar. Operaciones tiene un peso del 40 % bajo Speed-First, no el 33,3 % implícito de un promedio sin ponderar.

| Método | CX (3.8) | Ops (2.9) | Rev (3.2) | Compuesta | Calificación |
|---|---|---|---|---|---|
| Promedio sin ponderar | 33.3% | 33.3% | 33.3% | **3.30** | B |
| Pesos Speed-First | 30% | **40%** | 30% | **3.20** | B |
| Pesos Space-First | **50%** | 30% | 20% | **3.42** | B |

La calificación se mantiene en este caso — pero en el límite de 3,0 (B/C), un desplazamiento de 0,10 marca la diferencia entre una decisión pospuesta y una revisión operativa inmediata. Más importante aún, la diferencia de peso indica *dónde* intervenir: los operadores Speed-First deben tratar una puntuación de 2,9 en Operaciones como el riesgo principal, no como una preocupación secundaria equilibrada con CX.

**Enfoque corregido.** Aplique los pesos por tipo de tienda antes de comparar puntuaciones compuestas entre locales. Una cafetería Space-First con las mismas puntuaciones brutas se sitúa en un nivel de riesgo distinto porque CX aporta 20 puntos porcentuales adicionales.

<details>
<summary>Fórmula</summary>

```excel
=SUMPRODUCT(
  CHOOSE(
    MATCH(Store_Type, {"Speed-First","Space-First"}, 0),
    Speed_Weights_Vector,
    Space_Weights_Vector
  ),
  Dimension_Scores_Vector
)
```
</details>

---

### Trampa 2 — La muestra de un solo turno infla la puntuación

**Se tomó una decisión.** Un auditor visita durante la hora punta matutina, observa un 4,5 en Eficiencia del flujo de trabajo del barista e informa de operaciones sólidas. El propietario decide que no se necesita ninguna intervención en el flujo de trabajo.

**El defecto inadvertido.** La hora punta matutina es el turno que exige más preparación. La dotación de personal en las horas de baja afluencia es más reducida. La misma métrica puntúa 2,0 a las 18:30 porque un solo barista cubre tanto la barra como las bebidas sin un protocolo de traspaso definido.

**Antes — auditoría de un solo turno:**

```
OPS_001 Morning Peak only:  4.5
Reported average:           4.5
Decision:                   No action required
```

**Después — auditoría de tres turnos:**

```
OPS_001 Morning Peak:    4.5
OPS_001 Afternoon Rush:  3.5
OPS_001 Off-Peak:        2.0
Three-shift average:     3.33
Red flag:                NOT triggered (3.33 ≥ 2.5 threshold)
Additional finding:      Off-Peak requires a targeted staffing review
```

**Por qué el razonamiento es incorrecto.** Las auditorías de un solo turno muestrean por defecto la ventana mejor preparada. Los niveles de personal en la hora punta matutina suelen ser más numerosos y previamente informados. El instrumento exige tres ventanas operativas precisamente para evitar que este sesgo de selección entre en la puntuación compuesta.

**Resultado corregido.** La fórmula AVERAGEIF excluye los turnos sin puntuar mediante una condición `>0`. Ningún dato de turno ausente distorsiona la compuesta hacia cero — pero la puntuación de las horas de baja afluencia por sí sola es un dato señalable para la sección de recomendaciones incluso cuando el promedio supera el umbral.

<details>
<summary>Fórmula</summary>

```excel
-- Per-metric, per-shift average across the three shift columns:
=IFERROR(AVERAGEIF(B_row:D_row, ">0"), "")

-- Excludes shifts where no score was entered (value = 0)
-- A shift left unscored does not count as a zero in the average
```
</details>

---

### Trampa 3 — La fuga de ingresos se mantiene cualitativa

**Se tomó una decisión.** Un auditor anota "algunos clientes se fueron por la longitud de la cola" en un informe narrativo. El propietario reconoce la observación y no toma medidas porque el impacto en dólares no está claro.

**El defecto inadvertido.** "Algunos clientes" no es algo sobre lo que se pueda actuar. Con 8 salidas observadas al día y un gasto medio de $5,50, la fuga mensual es de $1.320 — 2,6 veces el umbral de intervención de $500. La observación justificaba un despliegue inmediato de gestión de colas, no una nota narrativa al pie pospuesta para el próximo trimestre.

**Antes — cualitativa:**

```
Finding:        Queue management below standard
Evidence:       "Customers leaving during peak"
Recommendation: "Consider reviewing queue process"
Owner response: Noted for Q3 planning
```

**Después — cuantificada:**

```
Lost customers (observed, daily):  8
ASP:                               $5.50
Monthly leakage:                   8 × $5.50 × 30 = $1,320
Alert threshold exceeded:          YES (> $500)
Required decision:                 Queue intervention; ROI positive within ~2 weeks
```

**Enfoque corregido.** El auditor cuenta las deserciones observables durante la visita. La fórmula convierte ese recuento en una cifra mensual y aplica un umbral binario que fuerza una clasificación de decisión — no una recomendación narrativa que puede posponerse.

Tenga en cuenta la limitación inherente: el recuento es una observación solo de las deserciones visibles. Los clientes que hacen cola y esperan (experiencia degradada, probabilidad de retorno reducida) no se capturan. La cifra de fuga es, por tanto, una estimación mínima, no máxima.

<details>
<summary>Fórmula</summary>

```excel
Monthly_Revenue_Leakage = Lost_Customers_Daily × ASP × 30

Alert threshold:
=IF(Monthly_Revenue_Leakage > 500, "IMMEDIATE ACTION REQUIRED", "MONITOR")
```
</details>

---

## Escenario de ejemplo

**Local:** Bar de espresso urbano independiente, centrado en el servicio para llevar del centro de la ciudad.  
**Tipo de tienda:** Speed-First — Operaciones ponderada al 40 %, CX al 30 %, Ingresos al 30 %.  
**ASP:** $4,80 · **Clientes perdidos observados:** 6 al día.

**Puntuaciones brutas — hora punta matutina (solo el Turno 1, con fines ilustrativos):**

| Métrica | Puntuación | Anclaje de rúbrica |
|---|---|---|
| CX_001 Amabilidad del servicio | 4 | Cálido y comprometido; se observa un trato personal |
| CX_002 Cola y tiempo de espera | 3 | Espera de 4–6 min; sin comunicación proactiva del tiempo estimado |
| CX_003 Precisión del pedido | 5 | Cero errores confirmados en todos los pedidos |
| CX_004 Ambiente y asientos | 4 | Agradable; buena disponibilidad de asientos |
| OPS_001 Flujo de trabajo del barista | 4 | Eficiente; movimientos desperdiciados mínimos |
| OPS_002 Fiabilidad del equipamiento | 3 | Funcional; retrasos menores de mantenimiento |
| OPS_003 Limpieza de la barra | 4 | Limpia de forma constante durante todo el turno |
| OPS_004 Disponibilidad de stock | 3 | 1–2 artículos no disponibles; el personal informó a los clientes |
| REV_001 Visibilidad del menú | 3 | Menú visible; intentos de venta adicional ocasionales |
| REV_002 Capacidad en hora punta | **2** | Ocupación del 40–60%; sin estrategia promocional |
| REV_003 Percepción de precios | 4 | Valor comunicado con claridad; señales de gama alta presentes |
| REV_004 Gasto secundario | 3 | Mención ocasional de complementos por parte del personal |

**Resultados calculados:**

```
CX  dimension avg:  (4+3+5+4) / 4 = 4.00
Ops dimension avg:  (4+3+4+3) / 4 = 3.50
Rev dimension avg:  (3+2+4+3) / 4 = 3.00

Composite (Speed-First):
(4.00 × 0.30) + (3.50 × 0.40) + (3.00 × 0.30)
= 1.20 + 1.40 + 0.90 = 3.50

Performance Grade:   B
Monthly leakage:     6 × $4.80 × 30 = $864  →  ALERT TRIGGERED
Red flag:            REV_002  (score 2.0 < 2.5 threshold)
```

**Interpretación analítica.** La calificación B indica que la operación cumple los estándares básicos en términos generales. La fuga mensual de $864 y la señal de alerta de REV_002 son incompatibles con una postura de acción diferida. La subutilización de la capacidad (40–60 % en hora punta) es la causa inmediata de la fuga. La puntuación compuesta lo enmascara: con 3,50, el local parece estable; la señal de alerta extraída mediante FILTER pone al descubierto el punto de fallo específico.

**Recomendación.** Despliegue de un sistema de cola anticipada o de reservas ligeras, combinado con una promoción de afluencia en hora punta. No se requiere inversión de capital.

**Implicación de decisión.** Si REV_002 mejora de 2,0 a 3,5 tras la intervención, el promedio de la dimensión Ingresos sube de 3,00 a 3,38. La puntuación compuesta se desplaza de 3,50 a 3,61 — no es un cambio de calificación, pero sí un movimiento significativo hacia el límite del nivel A (4,0) que se vuelve rastreable en el próximo ciclo de auditoría.

**Nota de incertidumbre.** Esta proyección asume que los Turnos 2 y 3 producen puntuaciones equivalentes o mejores. Una auditoría completa de tres turnos puede revelar que la hora punta de la tarde rinde peor que la hora punta matutina en Operaciones, lo que atenuaría la puntuación compuesta. Las proyecciones de un solo turno deben tratarse como orientativas, no precisas.

---

## Lo que el libro de trabajo realmente ofrece

- Una **puntuación compuesta única y comparable** que tiene en cuenta el modelo operativo, no solo los promedios brutos — comparable entre visitas y entre locales
- Una **calificación de rendimiento** (S / A / B / C / F) que comunica el nivel en un solo carácter, utilizable en presentaciones para clientes, informes para el consejo de administración y conversaciones de renegociación de arrendamientos
- Una **cifra de fuga financiera** que convierte la observación en lenguaje presupuestario y activa un umbral de acción binario en $500/mes
- Una **lista de acciones generada automáticamente** que extrae solo las métricas que fallan — sin filtrado manual, sin elementos perdidos por no haber sido evidentes durante la reunión de cierre
- Un **mandato de cobertura de tres turnos** que elimina el sesgo de muestra única sin aumentar la complejidad de la visita
- Un **sistema de pesos por tipo de tienda** que evita que una barra para llevar se evalúe según los estándares de una cafetería de barrio
- Un **diccionario de rúbrica de 60 entradas** que alinea a dos auditores independientes que puntúan la misma cafetería con una diferencia de aproximadamente 0,5 puntos

---

## Lógica del libro de trabajo (para el revisor escéptico)

**Arquitectura de hojas — flujo de datos estricto en una sola dirección, sin referencias circulares, sin macros:**

```
Setup_Parameters  →  Input_Data  →  Calc_Engine [hidden]  →  Dashboard_Report
     (config)          (field obs)       (processing)            (output)
```

**Setup_Parameters** contiene: la matriz de pesos por tipo de tienda (Speed-First frente a Space-First en 3 dimensiones), el diccionario de rúbrica de 60 entradas con clave `Metric_ID_Score` (p. ej., `CX_001_3`), los objetivos de referencia de la industria (CX: 4,0, Ops: 3,8, Rev: 3,5) y la configuración de ASP.

**Input_Data** recopila: 12 puntuaciones de métricas × 3 turnos (36 celdas puntuadas), descripciones de rúbrica generadas por XLOOKUP que se completan automáticamente al introducir la puntuación, y una entrada financiera (clientes perdidos observados por día). Todas las celdas de puntuación están restringidas por listas desplegables de validación de datos a números enteros del 1 al 5, con información sobre herramientas de rúbrica en línea.

**Calc_Engine** (oculto en la entrega al cliente) ejecuta la cadena de procesamiento completa: promedios de dimensión por turno → promedios generales con AVERAGEIF → puntuación compuesta ponderada con SUMPRODUCT → clasificación de calificación con IFS → multiplicación de la fuga financiera → vector booleano de señales de alerta (TRUE donde el promedio < 2,5). Los resultados son consumidos por el panel mediante rangos con nombre.

**Dashboard_Report** presenta: tres fichas de indicadores clave (puntuación compuesta, calificación, fuga mensual con umbral de alerta de $500) → fichas de puntuación por dimensión con delta respecto a la referencia → panel de acciones impulsado por FILTER que lista solo las métricas con señal de alerta junto con soluciones preescritas → marcador de posición de gráfico de radar que compara los valores reales frente a la referencia en tres ejes. Área de impresión fijada en A1:H45, A4 vertical, ajustado a la página.

**Rangos con nombre registrados (14):** `Store_Type`, `ASP_Value`, `Speed_Weights`, `Space_Weights`, `Benchmark_Scores`, `Overall_Score`, `Performance_Grade`, `Monthly_Leakage`, `Metric_Report_Range`, `Red_Flag_Vector`, `CX_Score_Overall`, `Ops_Score_Overall`, `Rev_Score_Overall`, `Lost_Customers_Daily`.

---

## Notas de implementación

<details>
<summary>Fórmula de la puntuación compuesta ponderada</summary>

```excel
=SUMPRODUCT(
  CHOOSE(
    MATCH(Setup_Parameters!$B$4, {"Speed-First","Space-First"}, 0),
    Setup_Parameters!$B$8:$B$10,    -- Speed-First weights [0.30, 0.40, 0.30]
    Setup_Parameters!$C$8:$C$10     -- Space-First weights [0.50, 0.30, 0.20]
  ),
  IFERROR(E5:E7 * 1, 0)             -- CX / Ops / Rev overall averages from Calc_Engine
)
```
</details>

<details>
<summary>Calificación de rendimiento</summary>

```excel
=IFS(
  B11 >= 4.5, "S",
  B11 >= 4.0, "A",
  B11 >= 3.0, "B",
  B11 >= 2.0, "C",
  B11 <  2.0, "F"
)
```
</details>

<details>
<summary>Extracción de señales de alerta — fórmula FILTER del panel</summary>

```excel
=IFERROR(
  FILTER(
    Calc_Engine!$A$21:$E$32,           -- Metric_Report_Range: ID, Dim, Name, Score, Action
    Calc_Engine!$F$21:$F$32 = TRUE,    -- Red_Flag_Vector: TRUE where avg < 2.5
    "No high-risk issues identified. All metrics meet the 2.5 threshold."
  ),
  "–"
)
```
</details>

<details>
<summary>Descripción de rúbrica con XLOOKUP — visualización de criterios en vivo de Input_Data</summary>

```excel
=IFERROR(
  XLOOKUP(
    A{row} & "_" & TEXT(D{row}, "0"),      -- Key: e.g. "CX_001_3"
    Setup_Parameters!$A$35:$A$94,           -- Key column (60 entries)
    Setup_Parameters!$B$35:$B$94,           -- Criteria description column
    "Select a score (1–5) to display criteria"
  ),
  ""
)
```
</details>

<details>
<summary>Gráfico de radar — construcción manual del grupo B</summary>

El libro de trabajo contiene un marcador de posición gris en las filas 25–44 del panel, con instrucciones de construcción completas. El gráfico no puede generarse de forma programada mediante openpyxl.

```
Chart type:     Radar (Filled)
Series 1:       Calc_Engine!$E$5:$E$7        — Actual scores
Series 2:       Setup_Parameters!$B$14:$B$16  — Benchmarks
Axis labels:    Customer Experience | Operations | Revenue Opportunity

Series 1 fill:  #2251FF at 30% transparency, border #2251FF 1.5pt solid
Series 2 fill:  none, border #888888 1pt dashed
Background:     transparent | No chart border | Legend: bottom centre
Resize to:      A25:H44
```
</details>

<details>
<summary>Reglas de validación y lista de verificación de entrega</summary>

```
Score cells:     integer list {1,2,3,4,5}, tooltip explains each level
Store Type:      list {"Speed-First","Space-First"}, error alert on invalid entry
All formulas:    wrapped in IFERROR(..., "") — no #DIV/0! or #N/A on Dashboard
Calc_Engine:     set sheet_state = hidden before client delivery
Print area:      Dashboard_Report!$A$1:$H$45, A4 portrait, fit-to-width 1 page
Gridlines:       disabled on Dashboard_Report (View > Show Gridlines unchecked)
```
</details>

---

## Obtener el libro de trabajo

Descargue `cafe_audit_tool.xlsx` desde este repositorio.

Ábralo en Excel de escritorio, iOS, Android u Online. Sin macros. No se requiere ninguna configuración más allá de introducir el Tipo de Tienda (Setup_Parameters B4) y el ASP (Setup_Parameters E4).

El archivo de simulación `cafe_audit_demo.html` refleja toda la lógica de Calc_Engine en el navegador. Ábralo localmente para validar los resultados de las fórmulas frente al libro de trabajo de Excel antes de la entrega al cliente.

---

## Limitaciones

**Lo que este libro de trabajo no puede hacer:**

- No captura la experiencia reportada por los clientes. Las puntuaciones reflejan únicamente la observación del auditor. Una cafetería puede puntuar 4,5 en CX mientras los clientes reportan 2,8 en NPS — la observación y la percepción no son métricas equivalentes.
- No realiza un seguimiento de datos longitudinales. Cada auditoría es una instantánea de un solo punto. El análisis de tendencias requiere el registro manual de las puntuaciones compuestas y las cifras de fuga entre visitas en un rastreador independiente.
- El cálculo de la fuga financiera asume que todos los clientes perdidos tienen el mismo ASP. Si los clientes que se marchan se inclinan hacia un gasto alto, la estimación es conservadora. Si se inclinan hacia un gasto bajo (p. ej., visitantes que solo miran en lugar de compradores), la fuga se sobreestima.
- Los valores de referencia (CX: 4,0, Ops: 3,8, Rev: 3,5) y el umbral de alerta de $500 son valores predeterminados configurables, no promedios de la industria validados estadísticamente. Los usuarios en mercados urbanos de alto volumen o contextos suburbanos de bajo volumen deben recalibrar Setup_Parameters!B14:B16 y el umbral de fuga de forma independiente.
- El umbral de señal de alerta de 2,5 es un corte de mitad de escala. No pondrá al descubierto métricas que puntúen 2,6 y que pueden seguir representando un riesgo operativo según el contexto.

---

## Otras herramientas de decisión que podría necesitar

| Escenario | Instrumento sugerido |
|---|---|
| Comparación de carteras de múltiples locales (10+ cafeterías) | Agregue las puntuaciones de auditoría en un rastreador de Excel independiente usando GROUPBY/PIVOTBY sobre las columnas de puntuación compuesta y fuga |
| Análisis de la brecha entre la percepción del cliente y la puntuación operativa | Combine la puntuación compuesta de la auditoría con datos de NPS o CSAT; análisis de correlación en un modelo independiente |
| Medición del retorno de la inversión de la formación del personal | Auditoría antes/después usando este libro de trabajo; el delta en la puntuación compuesta y la fuga mensual es el indicador clave principal |
| Soporte de decisión para la renovación del contrato de arrendamiento | Combine la puntuación compuesta con los ingresos por metro cuadrado y el recuento de afluencia en un modelo financiero independiente |
| Análisis de precios del menú | Requiere datos de TPV a nivel de transacción; fuera del alcance de un instrumento de auditoría basado en la observación |
| Comparación competitiva frente a marcas determinadas | Requiere datos de referencia de origen externo; las referencias del libro de trabajo son configurables pero no se actualizan automáticamente |

---

## Licencia

Copyright 2025. Publicado bajo la [Apache License 2.0](https://www.apache.org/licenses/LICENSE-2.0).

Use, modifique y distribuya este libro de trabajo y los archivos asociados bajo los términos de la Apache License 2.0. Se requiere atribución. No se ofrece ninguna garantía, expresa ni implícita. Consulte el archivo `LICENSE` para conocer los términos completos.

---

*Instrumento de soporte de decisión. No es un sistema empresarial.*

---

## Auditoría de cafetería · Motor de diagnóstico 3D

[![Licencia](https://img.shields.io/badge/许可证-Apache_2.0-blue.svg)](https://opensource.org/licenses/Apache-2.0)
[![Plataforma](https://img.shields.io/badge/平台-Excel%20%7C%20Web-navy)](https://github.com)
[![Tipo de herramienta](https://img.shields.io/badge/工具类型-决策支持-2251FF)](https://github.com)

**Un marco de auditoría estructurado que convierte la observación in situ de la cafetería en una puntuación diagnóstica ponderada, una estimación de la fuga de ingresos y una lista de acciones priorizada — en menos de 30 minutos por visita.**

Esta versión en español es coherente con la versión en inglés. Todas las fórmulas, la lógica de ponderación, las notas de arquitectura y las cláusulas de limitación se aplican igualmente. Los detalles técnicos completos se encuentran en las secciones desplegables de la versión en inglés.

[▶ Demostración en línea](cafe_audit_demo.html) · [⬇ Descargar libro de trabajo](cafe_audit_tool.xlsx)
