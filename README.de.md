[ 🌐 عربي ](README.ar.md) | [ 🇩🇪 Deutsch ](README.de.md) | [ 🇫🇷 Français ](README.fr.md) | [ 🇳🇱 Nederlands ](README.nl.md) | [ 🇪🇸 Español ](README.sp.md) | [ 🇬🇧 English ](README.md)

# Café-Audit · 3D-Diagnose-Engine

[![Lizenz](https://img.shields.io/badge/License-Apache_2.0-blue.svg)](https://opensource.org/licenses/Apache-2.0)
[![Plattform](https://img.shields.io/badge/Platform-Excel%20%7C%20Web-navy)](https://github.com)
[![Werkzeugtyp](https://img.shields.io/badge/Tool_Type-Decision_Support-2251FF)](https://github.com)

**Ein strukturiertes Audit-Framework, das Vor-Ort-Beobachtungen im Café in eine gewichtete Diagnosebewertung, eine Einschätzung von Einnahmenverlusten und eine priorisierte Maßnahmenliste umwandelt — in unter 30 Minuten pro Besuch.**

---

[▶ Live-Demo](https://hyvoid.github.io/Caf-audit-tool-for-efficient-reviews/) · [⬇ Arbeitsmappe herunterladen](cafe_audit_tool.xlsx) · 

---

 <img width="1536" height="1024" alt="ChatGPT Image Jul 14, 2026, 06_20_33 PM" src="https://github.com/user-attachments/assets/99030acd-62a9-4f24-b7dd-425c643c8a39" />


---

## Welche Entscheidung unterstützt dieses Werkzeug?

1. **Welche operative Dimension die meisten Einnahmen kostet** — Kundenerfahrung, Betrieb oder Umsatzchance — gewichtet nach Geschäftstyp?
2. **Liegt dieses Café über oder unter dem Mindestschwellenwert der Branche** — und um wie viel, ausgedrückt als Gesamtpunktzahl auf einer verifizierten 5-Punkte-Skala?
3. **Wie viele Einnahmen monatlich verloren gehen** durch beobachtbare Mängel, und übersteigt diese Zahl den Eingriffsschwellenwert von 500 $?
4. **Welche konkreten Kennzahlen sofortiges Handeln erfordern** — und was ist der erste Korrekturschritt für jede einzelne?

---

## Über den Entwickler

Drei operative Rahmenbedingungen bestimmen jede Gestaltungsentscheidung in dieser Arbeitsmappe.

**Entscheidungsunterstützung statt Systemersatz.** Dies ist keine POS-Integration, keine Loyalty-Analyseplattform und keine Betriebsmanagement-Suite. Es ist ein strukturiertes Beobachtungsinstrument, das aus im Feld auf einem Tablet in einem einzigen Besuch erhobenen Daten ein entscheidungsreifes Ergebnis erzeugt.

**Produktisiertes Denken statt Funktionsanhäufung.** Jede Formel dient einer Entscheidung. Die SUMPRODUCT-Gewichtungs-Engine existiert, weil ungewichtete Durchschnittswerte die Leistung von Cafés mit unterschiedlichen Betriebsmodellen systematisch falsch darstellen. Die FILTER-basierte Maßnahmenliste existiert, weil ein Dashboard ohne Priorisierung weiterhin menschliche Urteilskraft zum Handeln erfordert.

**Reibungsarme Ausführung statt Umsetzungskomplexität.** Keine Makros. Keine VBA-Abhängigkeiten. Keine externen Datenverbindungen. Die Arbeitsmappe läuft unverändert auf Excel für iOS, Android, iPadOS und Excel Online.

---

## Für wen dieses Werkzeug gedacht ist

| Benutzer | Typische Situation | Unterstützte Entscheidung | Warum bestehende Methoden scheitern |
|---|---|---|---|
| Unabhängiger Café-Berater | Audits von 4–12 Standorten pro Monat für einen Portfoliobesitzer | Welche Standorte Kapital brauchen, welche Prozesse, welche Personal | Besuche nach Bauchgefühl erzeugen inkonsistente Ergebnisse; Kunden können nicht über Besuche hinweg vergleichen |
| Betriebsleiter Gastronomie (F&B) | Vergleich der Filialleistung für den Quartalsreview | Wo der Regionalleiter seinen Fokus setzen sollte | Verkaufsdaten auf Filialebene isolieren die Grundursache nicht; zwei Cafés mit identischem Umsatz können eine entgegengesetzte Betriebsgesundheit haben |
| Café-Inhaber / Betreiber | Selbstdiagnose vor einer Vertragsverlängerung oder Umbauentscheidung | Ob das Problem beim Produkt, beim Personal oder beim Raum liegt | Die Nähe des Inhabers erzeugt blinde Flecken; kein externer Bezugspunkt für Bewertungen |
| Anbieter von Gastronomieschulungen | Bewertung vor/nach der Schulung für Kunden-Cafés | Ob die Schulung eine messbare operative Verbesserung erzielt hat | Der Vorher-Nachher-Vergleich erfordert ein konsistentes Messinstrument, das die meisten Betreiber nicht haben |

---

## Warum die meisten Café-Audit-Fehler keine Urteilsfehler sind

Der Prüfer verfügt in der Regel über ausreichend Expertise. Der Fehler entsteht früher — in der Messstruktur.

**Wo der Prozess bricht:**

```
Unstructured observation
    → inconsistent metric coverage per visit
    → no rubric → scores reflect auditor mood, not café performance
    → averages across dimensions hide the worst-performing metric
    → action list ranked by salience, not by weighted impact
    → owner receives a report; no decision follows
```

**Mit einem strukturierten Instrument:**

```
Rubric-anchored 5-point scoring (12 metrics × 3 shifts)
    → AVERAGE per metric across shifts → removes single-event bias
    → SUMPRODUCT × store-type weights → composite reflects operating model
    → IFS grade → single comparable output
    → FILTER(Red_Flag_Vector = TRUE) → action list auto-extracts
    → Lost customers × ASP × 30 → converts observation to financial stakes
```

**Operative Implikation:** Ein Prüfer, der „Warteschlange & Wartezeit“ intuitiv mit 3 bewertet, bewertet sie möglicherweise mit 2, wenn er gezwungen ist zu antworten: *"Is the wait 4–6 minutes with no proactive ETA communication?"* Das Bewertungsraster beseitigt die Lücke zwischen zwei Prüfern, die dieselbe Warteschlange beobachten.

Die Arbeitsmappe fällt das Urteil nicht. Sie verhindert, dass ein Urteil ohne Definition gefällt wird.

---

## Drei Fallen, in die auch erfahrene Betreiber tappen

### Falle 1 — Der ungewichtete Durchschnitt verfälscht das Betriebsmodell

**Eine Entscheidung wurde getroffen.** Ein Café erzielt 3,8 in Kundenerfahrung und 2,9 im Betrieb. Der ungewichtete Durchschnitt über drei Dimensionen beträgt 3,30. Der Betreiber schlussfolgert, die Leistung sei akzeptabel, und verschiebt den Eingriff.

**Der unbemerkte Fehler.** Das Café ist ein auf Take-away fokussierter Speed-First-Betrieb. Der Betrieb trägt unter Speed-First ein Gewicht von 40 %, nicht die impliziten 33,3 % eines ungewichteten Durchschnitts.

| Methode | CX (3.8) | Ops (2.9) | Rev (3.2) | Gesamtpunktzahl | Note |
|---|---|---|---|---|---|
| Ungewichteter Durchschnitt | 33.3% | 33.3% | 33.3% | **3.30** | B |
| Speed-First-Gewichte | 30% | **40%** | 30% | **3.20** | B |
| Space-First-Gewichte | **50%** | 30% | 20% | **3.42** | B |

Die Note hält hier stand — aber an der 3,0-Grenze (B/C) ist eine Verschiebung um 0,10 der Unterschied zwischen einer aufgeschobenen Entscheidung und einer sofortigen Betriebsüberprüfung. Wichtiger noch: Der Gewichtsunterschied signalisiert *wo* eingegriffen werden muss: Speed-First-Betreiber müssen einen Betriebswert von 2,9 als primäres Risiko behandeln, nicht als sekundäre Sorge, die gegen CX abgewogen wird.

**Korrigierter Ansatz.** Wenden Sie Geschäftstyp-Gewichte an, bevor Sie Gesamtpunktzahlen über Standorte hinweg vergleichen. Ein Space-First-Café mit identischen Rohwerten liegt in einer anderen Risikostufe, weil CX 20 zusätzliche Prozentpunkte trägt.

<details>
<summary>Formel</summary>

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

### Falle 2 — Die Einzelschicht-Stichprobe überhöht die Bewertung

**Eine Entscheidung wurde getroffen.** Ein Prüfer besucht während der Morgenspitzenzeit, beobachtet 4,5 bei der Barista-Workflow-Effizienz und meldet einen starken Betrieb. Der Inhaber entscheidet, dass kein Workflow-Eingriff erforderlich ist.

**Der unbemerkte Fehler.** Die Morgenspitzenzeit ist die am besten vorbereitete Schicht. In den Nebenzeiten ist das Personal knapper besetzt. Dieselbe Kennzahl erzielt 2,0 um 18:30 Uhr, weil ein Barista ohne definiertes Übergabeprotokoll sowohl die Theke als auch die Getränke abdeckt.

**Vorher — Einzelschicht-Audit:**

```
OPS_001 Morning Peak only:  4.5
Reported average:           4.5
Decision:                   No action required
```

**Nachher — Dreischicht-Audit:**

```
OPS_001 Morning Peak:    4.5
OPS_001 Afternoon Rush:  3.5
OPS_001 Off-Peak:        2.0
Three-shift average:     3.33
Red flag:                NOT triggered (3.33 ≥ 2.5 threshold)
Additional finding:      Off-Peak requires a targeted staffing review
```

**Warum die Argumentation falsch ist.** Einzelschicht-Audits erfassen standardmäßig das am besten vorbereitete Zeitfenster. Die Personalstärke in der Morgenspitzenzeit ist typischerweise höher und im Voraus eingewiesen. Das Instrument schreibt drei operative Zeitfenster vor, genau um zu verhindern, dass dieser Auswahlfehler in die Gesamtpunktzahl eingeht.

**Korrigiertes Ergebnis.** Die AVERAGEIF-Formel schließt unbewertete Schichten mithilfe einer Bedingung `>0` aus. Fehlende Schichtdaten verzerren die Gesamtpunktzahl nicht gegen null — aber der Nebenzeiten-Wert allein ist ein Datenpunkt, der eine rote Flagge setzt, selbst wenn der Durchschnitt den Schwellenwert übersteigt.

<details>
<summary>Formel</summary>

```excel
-- Per-metric, per-shift average across the three shift columns:
=IFERROR(AVERAGEIF(B_row:D_row, ">0"), "")

-- Excludes shifts where no score was entered (value = 0)
-- A shift left unscored does not count as a zero in the average
```
</details>

---

### Falle 3 — Einnahmenverluste bleiben qualitativ

**Eine Entscheidung wurde getroffen.** Ein Prüfer vermerkt in einem narrativen Bericht „einige Kunden sind wegen der Warteschlange gegangen“. Der Inhaber erkennt die Beobachtung an und ergreift keine Maßnahme, da die Auswirkung in Dollar unklar ist.

**Der unbemerkte Fehler.** „Einige Kunden“ ist nicht handlungsrelevant. Bei 8 beobachteten Abgängen pro Tag mit durchschnittlichen Ausgaben von 5,50 $ beträgt der monatliche Einnahmenverlust 1.320 $ — das 2,6-Fache des Eingriffsschwellenwerts von 500 $. Die Beobachtung hätte einen sofortigen Einsatz von Warteschlangenmanagement gerechtfertigt, keine narrative Fußnote, die auf das nächste Quartal verschoben wird.

**Vorher — qualitativ:**

```
Finding:        Queue management below standard
Evidence:       "Customers leaving during peak"
Recommendation: "Consider reviewing queue process"
Owner response: Noted for Q3 planning
```

**Nachher — quantifiziert:**

```
Lost customers (observed, daily):  8
ASP:                               $5.50
Monthly leakage:                   8 × $5.50 × 30 = $1,320
Alert threshold exceeded:          YES (> $500)
Required decision:                 Queue intervention; ROI positive within ~2 weeks
```

**Korrigierter Ansatz.** Der Prüfer zählt beobachtbare Weggeher während des Besuchs. Die Formel wandelt diese Zahl in eine monatliche Größe um und wendet einen binären Schwellenwert an, der eine Entscheidungsklassifizierung erzwingt — keine narrative Empfehlung, die verschoben werden kann.

Beachten Sie die inhärente Einschränkung: Die Zahl ist eine Beobachtung ausschließlich sichtbarer Abgänge. Kunden, die sich anstellen und warten (beeinträchtigte Erfahrung, verringerte Rückkehrwahrscheinlichkeit), werden nicht erfasst. Die Verlustzahl ist daher eine Untergrenze, keine Obergrenze.

<details>
<summary>Formel</summary>

```excel
Monthly_Revenue_Leakage = Lost_Customers_Daily × ASP × 30

Alert threshold:
=IF(Monthly_Revenue_Leakage > 500, "IMMEDIATE ACTION REQUIRED", "MONITOR")
```
</details>

---

## Beispielszenario

**Standort:** Unabhängige urbane Espressobar, Innenstadtlage mit Take-away-Fokus.  
**Geschäftstyp:** Speed-First — Betrieb mit 40 % gewichtet, CX 30 %, Umsatz 30 %.  
**ASP:** 4,80 $ · **Beobachtete verlorene Kunden:** 6 pro Tag.

**Rohwerte — Morgenspitzenzeit (nur Schicht 1 zur Veranschaulichung):**

| Kennzahl | Wert | Anker im Bewertungsraster |
|---|---|---|
| CX_001 Freundlichkeit des Service | 4 | Warm und engagiert; persönliche Note beobachtet |
| CX_002 Warteschlange & Wartezeit | 3 | Wartezeit 4–6 min; keine proaktive ETA-Kommunikation |
| CX_003 Bestellgenauigkeit | 5 | Null Fehler über alle Bestellungen hinweg bestätigt |
| CX_004 Atmosphäre & Sitzgelegenheiten | 4 | Angenehm; gute Verfügbarkeit an Sitzplätzen |
| OPS_001 Barista-Workflow | 4 | Effizient; minimale überflüssige Bewegung |
| OPS_002 Zuverlässigkeit der Geräte | 3 | Funktional; kleinere Wartungsverzögerungen |
| OPS_003 Sauberkeit der Theke | 4 | Durchgehend sauber über die gesamte Schicht |
| OPS_004 Warenverfügbarkeit | 3 | 1–2 Artikel nicht verfügbar; Personal informierte die Kunden |
| REV_001 Sichtbarkeit der Speisekarte | 3 | Speisekarte sichtbar; gelegentliche Versuche zum Zusatzverkauf |
| REV_002 Spitzenkapazität | **2** | 40–60% belegt; keine Werbestrategie |
| REV_003 Preiswahrnehmung | 4 | Wert klar kommuniziert; Premium-Signale vorhanden |
| REV_004 Zusatzausgaben | 3 | Gelegentliche Erwähnung von Zusatzangeboten durch das Personal |

**Berechnete Ergebnisse:**

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

**Analytische Interpretation.** Note B zeigt, dass der Betrieb insgesamt die Basisanforderungen erfüllt. Der monatliche Einnahmenverlust von 864 $ und die rote Flagge bei REV_002 passen nicht zu einer Haltung des Aufschiebens. Die Kapazitätsunterauslastung (40–60 % in der Spitzenzeit) ist die unmittelbare Ursache des Verlusts. Die Gesamtpunktzahl verschleiert dies: Bei 3,50 sieht der Standort stabil aus; die per FILTER extrahierte rote Flagge deckt den konkreten Schwachpunkt auf.

**Empfehlung.** Einsatz einer virtuellen Warteschlange oder eines leichten Reservierungssystems, kombiniert mit einer Aktion zur Steigerung der Kundenfrequenz in Spitzenzeiten. Keine Investitionsausgaben erforderlich.

**Implikation für die Entscheidung.** Wenn sich REV_002 nach dem Eingriff von 2,0 auf 3,5 verbessert, steigt der Durchschnitt der Umsatzdimension von 3,00 auf 3,38. Die Gesamtpunktzahl verschiebt sich von 3,50 auf 3,61 — keine Notenänderung, aber eine bedeutsame Bewegung in Richtung der A-Grenze (4,0), die im nächsten Audit-Zyklus verfolgbar wird.

**Hinweis zur Unsicherheit.** Diese Projektion setzt voraus, dass die Schichten 2 und 3 gleichwertige oder bessere Werte liefern. Ein vollständiges Dreischicht-Audit könnte zeigen, dass die Nachmittagsspitzenzeit in der Betriebsdimension schlechter abschneidet als die Morgenspitzenzeit, was die Gesamtpunktzahl dämpfen würde. Einzelschicht-Projektionen sollten als richtungsweisend, nicht als präzise behandelt werden.

---

## Was die Arbeitsmappe tatsächlich liefert

- Eine **einzige vergleichbare Gesamtpunktzahl**, die das Betriebsmodell berücksichtigt und nicht nur Rohdurchschnitte — vergleichbar über Besuche und Standorte hinweg
- Eine **Leistungsnote** (S / A / B / C / F), die die Stufe in einem einzigen Zeichen kommuniziert und in Kundenpräsentationen, Vorstandsunterlagen und Gesprächen zur Vertragsneuverhandlung einsetzbar ist
- Eine **Einnahmenverlustzahl**, die Beobachtung in Budgetsprache übersetzt und bei 500 $/Monat einen binären Aktionsschwellenwert auslöst
- Eine **automatisch generierte Maßnahmenliste**, die nur fehlgeschlagene Kennzahlen extrahiert — keine manuelle Filterung, keine Punkte übersehen, weil sie in der Nachbesprechung nicht hervorstachen
- Eine **Dreischicht-Abdeckungspflicht**, die Einzelstichproben-Verzerrungen beseitigt, ohne die Besuchskomplexität zu erhöhen
- Ein **Geschäftstyp-Gewichtungssystem**, das verhindert, dass eine Take-away-Bar nach den Standards eines Stadtteilcafés bewertet wird
- Ein **Bewertungsraster-Wörterbuch mit 60 Einträgen**, das zwei unabhängige Prüfer, die dasselbe Café bewerten, auf etwa 0,5 Punkte angleicht

---

## Logik der Arbeitsmappe (für den skeptischen Prüfer)

**Blattarchitektur — strikt einseitiger Datenfluss, keine Zirkelbezüge, keine Makros:**

```
Setup_Parameters  →  Input_Data  →  Calc_Engine [hidden]  →  Dashboard_Report
     (config)          (field obs)       (processing)            (output)
```

**Setup_Parameters** enthält: eine Gewichtungsmatrix nach Geschäftstyp (Speed-First vs. Space-First über 3 Dimensionen), ein Bewertungsraster-Wörterbuch mit 60 Einträgen, geschlüsselt als `Metric_ID_Score` (z. B. `CX_001_3`), Branchen-Benchmark-Ziele (CX: 4,0, Ops: 3,8, Rev: 3,5) sowie die ASP-Konfiguration.

**Input_Data** erhebt: 12 Kennzahlwerte × 3 Schichten (36 bewertete Zellen), XLOOKUP-gesteuerte Bewertungsraster-Beschreibungen, die bei Werteingabe automatisch gefüllt werden, sowie eine finanzielle Eingabe (beobachtete verlorene Kunden pro Tag). Alle Wertezellen sind über Datenvalidierungs-Dropdowns auf ganze Zahlen von 1–5 beschränkt, mit integrierten Bewertungsraster-Tooltips.

**Calc_Engine** (bei Kundenauslieferung ausgeblendet) führt die vollständige Verarbeitungskette aus: Dimensionsdurchschnitte pro Schicht → AVERAGEIF-Gesamtdurchschnitte → SUMPRODUCT-gewichtete Gesamtpunktzahl → IFS-Notenklassifizierung → Multiplikation des Einnahmenverlusts → boolescher Vektor roter Flaggen (TRUE, wo der Durchschnitt < 2,5). Die Ergebnisse werden vom Dashboard über benannte Bereiche bezogen.

**Dashboard_Report** stellt dar: drei KPI-Karten (Gesamtpunktzahl, Note, monatlicher Einnahmenverlust mit Warnschwellenwert 500 $) → Dimensionswert-Pillen mit Benchmark-Delta → FILTER-gesteuertes Aktionspanel, das nur Kennzahlen mit roter Flagge und vorgefertigte Abhilfen auflistet → Platzhalter für ein Radardiagramm, das Ist-Wert vs. Benchmark über drei Achsen vergleicht. Der Druckbereich ist fest auf A1:H45, A4 Hochformat und Seitenanpassung eingestellt.

**Registrierte benannte Bereiche (14):** `Store_Type`, `ASP_Value`, `Speed_Weights`, `Space_Weights`, `Benchmark_Scores`, `Overall_Score`, `Performance_Grade`, `Monthly_Leakage`, `Metric_Report_Range`, `Red_Flag_Vector`, `CX_Score_Overall`, `Ops_Score_Overall`, `Rev_Score_Overall`, `Lost_Customers_Daily`.

---

## Hinweise zur Implementierung

<details>
<summary>Formel für die gewichtete Gesamtpunktzahl</summary>

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
<summary>Leistungsnote</summary>

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
<summary>Extraktion roter Flaggen — Dashboard-FILTER-Formel</summary>

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
<summary>XLOOKUP-Bewertungsraster-Beschreibung — Live-Kriterienanzeige in Input_Data</summary>

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
<summary>Radardiagramm — manueller Aufbau (Kategorie B)</summary>

Die Arbeitsmappe enthält in den Dashboard-Zeilen 25–44 einen grauen Platzhalter mit vollständigen Aufbauanweisungen. Das Diagramm kann nicht programmatisch über openpyxl generiert werden.

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
<summary>Validierungsregeln und Auslieferungs-Checkliste</summary>

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

## Arbeitsmappe beschaffen

Laden Sie `cafe_audit_tool.xlsx` aus diesem Repository herunter.

Öffnen Sie sie in Excel Desktop, iOS, Android oder Online. Keine Makros. Keine Einrichtung erforderlich, abgesehen von der Eingabe des Geschäftstyps (Setup_Parameters B4) und des ASP (Setup_Parameters E4).

Die Simulationsdatei `cafe_audit_demo.html` spiegelt die gesamte Calc_Engine-Logik im Browser wider. Öffnen Sie sie lokal, um die Formelergebnisse vor der Kundenauslieferung gegen die Excel-Arbeitsmappe zu prüfen.

---

## Einschränkungen

**Was diese Arbeitsmappe nicht kann:**

- Sie erfasst keine von Kunden berichtete Erfahrung. Werte spiegeln ausschließlich die Beobachtung des Prüfers wider. Ein Café kann 4,5 in CX erzielen, während Kunden 2,8 beim NPS berichten — Beobachtung und Wahrnehmung sind keine gleichwertigen Kennzahlen.
- Sie verfolgt keine Längsschnittdaten. Jedes Audit ist eine Momentaufnahme zu einem Zeitpunkt. Trendanalysen erfordern die manuelle Protokollierung von Gesamtpunktzahlen und Verlustzahlen über Besuche hinweg in einem separaten Tracker.
- Die Berechnung des Einnahmenverlusts geht davon aus, dass alle verlorenen Kunden denselben ASP aufweisen. Wenn abwandernde Kunden zu hohen Ausgaben neigen, ist die Schätzung konservativ. Wenn sie zu niedrigen Ausgaben neigen (z. B. eher Interessenten als Käufer), überzeichnet sie den Verlust.
- Die Benchmark-Werte (CX: 4,0, Ops: 3,8, Rev: 3,5) und der Warnschwellenwert von 500 $ sind konfigurierbare Standardwerte, keine statistisch validierten Branchendurchschnitte. Nutzer in städtischen Märkten mit hohem Volumen oder in Vororten mit geringem Volumen müssen Setup_Parameters!B14:B16 und den Verlustschwellenwert unabhängig neu kalibrieren.
- Der Schwellenwert von 2,5 für rote Flaggen ist ein Schnitt in der Skalenmitte. Er wird Kennzahlen mit einem Wert von 2,6 nicht erfassen, die im Kontext dennoch ein operationelles Risiko darstellen können.

---

## Weitere Entscheidungswerkzeuge, die Sie benötigen könnten

| Szenario | Empfohlenes Instrument |
|---|---|
| Vergleich von Multi-Standort-Portfolios (10+ Cafés) | Auditwerte in einem separaten Excel-Tracker mit GROUPBY/PIVOTBY über die Gesamtpunktzahl- und Verlustspalten aggregieren |
| Analyse der Lücke zwischen Kundenwahrnehmung und Betriebswert | Audit-Gesamtpunktzahl mit NPS- oder CSAT-Daten paaren; Korrelationsanalyse in einem separaten Modell |
| Messung des ROI von Personalschulungen | Audit vor/nach mit dieser Arbeitsmappe; die Differenz von Gesamtpunktzahl und monatlichem Verlust ist die primäre KPI |
| Unterstützung bei Vertragsverlängerungsentscheidungen | Gesamtpunktzahl mit Umsatz pro Quadratmeter und Kundenfrequenz in einem separaten Finanzmodell kombinieren |
| Analyse der Speisekarten-Preisgestaltung | Erfordert transaktionsgenaue POS-Daten; außerhalb des Rahmens eines beobachtungsbasierten Audit-Instruments |
| Wettbewerbsvergleich mit benannten Marken | Erfordert extern bezogene Benchmark-Daten; die Benchmarks der Arbeitsmappe sind konfigurierbar, aber nicht automatisch aktualisiert |

---

## Lizenz

Copyright 2025. Veröffentlicht unter der [Apache License 2.0](https://www.apache.org/licenses/LICENSE-2.0).

Nutzung, Veränderung und Verbreitung dieser Arbeitsmappe und der zugehörigen Dateien unter den Bedingungen der Apache License 2.0. Namensnennung erforderlich. Es wird keine Gewährleistung gegeben, weder ausdrücklich noch stillschweigend. Die vollständigen Bedingungen finden Sie in der Datei `LICENSE`.

---

*Instrument zur Entscheidungsunterstützung. Kein Unternehmenssystem.*

---

## Café-Audit · 3D-Diagnose-Engine

[![Lizenz](https://img.shields.io/badge/许可证-Apache_2.0-blue.svg)](https://opensource.org/licenses/Apache-2.0)
[![Plattform](https://img.shields.io/badge/平台-Excel%20%7C%20Web-navy)](https://github.com)
[![Werkzeugtyp](https://img.shields.io/badge/工具类型-决策支持-2251FF)](https://github.com)

**Ein strukturiertes Audit-Framework, das Vor-Ort-Beobachtungen im Café in eine gewichtete Diagnosebewertung, eine Einschätzung von Einnahmenverlusten und eine priorisierte Maßnahmenliste umwandelt — in unter 30 Minuten pro Besuch.**

Diese deutschsprachige Version ist inhaltlich identisch mit der englischen Version. Alle Formeln, die Gewichtungslogik, die Architekturbeschreibung und die Einschränkungsklauseln gelten gleichermaßen. Die vollständigen technischen Details finden Sie in den aufklappbaren Bereichen der englischen Version.

[▶ Online-Demo](cafe_audit_demo.html) · [⬇ Arbeitsmappe herunterladen](cafe_audit_tool.xlsx)
