[ 🌐 عربي ](README.ar.md) | [ 🇩🇪 Deutsch ](README.de.md) | [ 🇫🇷 Français ](README.fr.md) | [ 🇳🇱 Nederlands ](README.nl.md) | [ 🇪🇸 Español ](README.sp.md) | [ 🇬🇧 English ](README.md)

# Café-audit · 3D-diagnostische motor

[![Licentie](https://img.shields.io/badge/License-Apache_2.0-blue.svg)](https://opensource.org/licenses/Apache-2.0)
[![Platform](https://img.shields.io/badge/Platform-Excel%20%7C%20Web-navy)](https://github.com)
[![Type tool](https://img.shields.io/badge/Tool_Type-Decision_Support-2251FF)](https://github.com)

**Een gestructureerd auditraamwerk dat observaties ter plaatse van een café omzet in een gewogen diagnostische score, een schatting van het omzetverlies en een geprioriteerde actielijst — in minder dan 30 minuten per bezoek.**

---

[▶ livedemo](https://hyvoid.github.io/Caf-audit-tool-for-efficient-reviews/) · [⬇ Download werkboek](cafe_audit_tool.xlsx) · 

---

 <img width="1536" height="1024" alt="ChatGPT Image Jul 14, 2026, 06_20_33 PM" src="https://github.com/user-attachments/assets/99030acd-62a9-4f24-b7dd-425c643c8a39" />


---

## Welke beslissing helpt dit u te nemen?

1. **Welke operationele dimensie de meeste omzet kost** — Klantbeleving, Operaties of Omzetkans — gewogen naar winkeltype?
2. **Presteert dit café boven of onder de minimale drempelwaarde van de sector** — en met hoeveel, uitgedrukt als samengestelde score op een geverifieerde 5-puntsschaal?
3. **Hoeveel omzetverlies er maandelijks ontstaat** door waarneembare falen, en overschrijdt dat bedrag de interventiedrempel van $500?
4. **Welke specifieke metrieken onmiddellijke actie vereisen** en wat is de eerste corrigerende stap voor elk ervan?

---

## Over de maker

Drie vaststaande beperkingen sturen elke ontwerpbeslissing in dit werkboek.

**Beslissingondersteuning boven systeemvervanging.** Dit is geen POS-integratie, geen loyaliteitsanalyseplatform en geen suite voor operationsmanagement. Het is een gestructureerd observatie-instrument dat uit veldgegevens — verzameld op een tablet tijdens één bezoek — een beslisklaare uitkomst produceert.

**Geproductiseerde redenering boven functie-accumulatie.** Elke formule dient een beslissing. De wegeningsengine met SUMPRODUCT bestaat omdat ongewegen gemiddelden prestaties voor café's met verschillende bedrijfsmodellen systematisch vertekenen. De actielijst op basis van FILTER bestaat omdat een dashboard zonder prioriteratie toch menselijk oordeel vereist om actie te ondernemen.

**Soepele uitvoering boven implementatiecomplexiteit.** Geen macro's. Geen VBA-afhankelijkheden. Geen externe dataverbindingen. Het werkboek draait ongewijzigd op Excel voor iOS, Android, iPadOS en Excel Online.

---

## Voor wie dit bedoeld is

| Gebruiker | Typische situatie | Ondersteunde beslissing | Waarom bestaande methoden falen |
|---|---|---|---|
| Onafhankelijke café-consultant | Audit van 4–12 locaties per maand voor een portefeuillehouder | Welke locaties kapitaal nodig hebben, welke procesaanpassing en welke personeelsinspanning | Bezoeken op buikgevoel leveren inconsistente resultaten op; klanten kunnen bezoeken niet met elkaar vergelijken |
| Operationsmanager horeca | Prestaties van vestigingen vergelijken voor de kwartaalevaluatie | Waar de aandacht van de regiomanager op te richten | Verkoopcijfers op vestigingsniveau isoleren de hoofdoorzaak niet; twee café's met identieke omzet kunnen een tegengestelde operationele gezondheid hebben |
| Caféeigenaar / exploitant | Zelfdiagnose voor een beslissing over huurverlenging of renovatie | Of het probleem ligt bij product, mensen of ruimte | Nabijheid van de eigenaar creëert blinde vlekken; geen extern referentiepunt voor scores |
| Aanbieder van horecatrainingen | Pre/post-beoordeling van trainingen voor café's van klanten | Of de training een meetbare operationele verbetering heeft opgeleverd | Een pre/post-vergelijking vereist een consistent meetinstrument, dat de meeste exploitanten niet hebben |

---

## Waarom de meeste fouten in café-audits geen beoordelingsfouten zijn

De auditor heeft meestal voldoende expertise. De fout ontstaat eerder — in de meetstructuur.

**Waar het proces strandt:**

```
Unstructured observation
    → inconsistent metric coverage per visit
    → no rubric → scores reflect auditor mood, not café performance
    → averages across dimensions hide the worst-performing metric
    → action list ranked by salience, not by weighted impact
    → owner receives a report; no decision follows
```

**Met een gestructureerd instrument:**

```
Rubric-anchored 5-point scoring (12 metrics × 3 shifts)
    → AVERAGE per metric across shifts → removes single-event bias
    → SUMPRODUCT × store-type weights → composite reflects operating model
    → IFS grade → single comparable output
    → FILTER(Red_Flag_Vector = TRUE) → action list auto-extracts
    → Lost customers × ASP × 30 → converts observation to financial stakes
```

**Operationele implicatie:** Een auditor die "Wachtrij & wachttijd" op intuïtie een 3 geeft, scoort het misschien 2 wanneer hij moet antwoorden op: *"Is the wait 4–6 minutes with no proactive ETA communication?"* De beoordelingsrubric neemt het verschil tussen twee auditors die dezelfde wachtrij observeren weg.

Het werkboek maakt de beoordeling niet. Het voorkomt dat een beoordeling wordt gegeven zonder definitie.

---

## Drie valkuilen die zelfs ervaren exploitanten vangen

### Valkuil 1 — Het ongewegen gemiddelde vertekent het bedrijfsmodel

**Er werd een beslissing genomen.** Een café scoort 3,8 op Klantbeleving en 2,9 op Operaties. Het ongewegen gemiddelde over drie dimensies is 3,30. De exploitant concludeert dat de prestaties acceptabel zijn en stelt interventie uit.

**Het onopgemerkte gebrek.** Het café is een afhaaluitgericht bedrijf van het type Speed-First. Operaties heeft onder Speed-First een gewicht van 40%, niet de geïmpliceerde 33,3% van een ongewegen gemiddelde.

| Methode | CX (3.8) | Ops (2.9) | Rev (3.2) | Samengestelde score | Cijfer |
|---|---|---|---|---|---|
| Ongewegen gemiddelde | 33.3% | 33.3% | 33.3% | **3.30** | B |
| Speed-First-gewichten | 30% | **40%** | 30% | **3.20** | B |
| Space-First-gewichten | **50%** | 30% | 20% | **3.42** | B |

Het cijfer blijft hier staan — maar op de grens van 3,0 (B/C) is een verschuiving van 0,10 het verschil tussen een uitgestelde beslissing en een onmiddellijke operationele evaluatie. Nog belangrijker: het gewichtsverschil geeft aan *waar* ingegrepen moet worden: exploitanten van het type Speed-First moeten een score van 2,9 op Operaties zien als het primaire risico, niet als een bijkomend zorgpunt tegenover CX.

**Gecorrigeerde aanpak.** Pas winkeltype-gewichten toe voordat u samengestelde scores tussen locaties vergelijkt. Een café van het type Space-First met identieke ruwe scores zit in een andere risicoklasse omdat CX er 20 extra procentpunten draagt.

<details>
<summary>Formule</summary>

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

### Valkuil 2 — Een enkele dienst als monster overdrijft de score

**Er werd een beslissing genomen.** Een auditor komt tijdens de ochtendpiek, ziet een 4,5 voor de efficiëntie van de baristaworkflow en rapporteert sterke prestaties. De eigenaar besluit dat geen interventie in de workflow nodig is.

**Het onopgemerkte gebrek.** De ochtendpiek is de dienst met de meeste voorbereiding. Buiten de spits is er minder personeel. Dezelfde metriek scoort 2,0 om 18:30 omdat één barista zowel toonbank als dranken dekt zonder een vastgelegd overdrachtsprotocol.

**Voor — audit over één dienst:**

```
OPS_001 Morning Peak only:  4.5
Reported average:           4.5
Decision:                   No action required
```

**Na — audit over drie diensten:**

```
OPS_001 Morning Peak:    4.5
OPS_001 Afternoon Rush:  3.5
OPS_001 Off-Peak:        2.0
Three-shift average:     3.33
Red flag:                NOT triggered (3.33 ≥ 2.5 threshold)
Additional finding:      Off-Peak requires a targeted staffing review
```

**Waarom de redenering onjuist is.** Audits over één dienst nemen standaard het best voorbereide moment als monster. Op de ochtendpiek is doorgaans meer personeel aanwezig en is dat vooraf ingelicht. Het instrument verplicht drie operationele vensters precies om te voorkomen dat deze selectiebias de samengestelde score binnendringt.

**Gecorrigeerd resultaat.** De AVERAGEIF-formule sluit ongescoorde diensten uit met een voorwaarde `>0`. Ontbrekende dienstgegevens vervormen de samengestelde score niet naar nul — maar de score buiten de spits op zichzelf rechtvaardigt al een waarschuwingssignaal in de sectie met aanbevelingen, ook wanneer het gemiddelde boven de drempelwaarde blijft.

<details>
<summary>Formule</summary>

```excel
-- Per-metric, per-shift average across the three shift columns:
=IFERROR(AVERAGEIF(B_row:D_row, ">0"), "")

-- Excludes shifts where no score was entered (value = 0)
-- A shift left unscored does not count as a zero in the average
```
</details>

---

### Valkuil 3 — Omzetverlies blijft kwalitatief

**Er werd een beslissing genomen.** Een auditor noteert "sommige klanten vertrokken vanwege de lengte van de wachtrij" in een narratief rapport. De eigenaar erkent de observatie en onderneemt geen actie omdat de gevolgen in dollars onduidelijk zijn.

**Het onopgemerkte gebrek.** "Enkele klanten" leidt niet tot actie. Bij 8 waargenomen vertrekken per dag met een gemiddelde besteding van $5,50 is het maandelijkse omzetverlies $1.320 — 2,6 keer de interventiedrempel van $500. De observatie rechtvaardigde onmiddellijke inzet van wachtrijbeheer, geen voetnoot in een rapport dat naar het volgende kwartaal wordt doorgeschoven.

**Voor — kwalitatief:**

```
Finding:        Queue management below standard
Evidence:       "Customers leaving during peak"
Recommendation: "Consider reviewing queue process"
Owner response: Noted for Q3 planning
```

**Na — gekwantificeerd:**

```
Lost customers (observed, daily):  8
ASP:                               $5.50
Monthly leakage:                   8 × $5.50 × 30 = $1,320
Alert threshold exceeded:          YES (> $500)
Required decision:                 Queue intervention; ROI positive within ~2 weeks
```

**Gecorrigeerde aanpak.** De auditor telt waarneembare vertrekkers tijdens het bezoek. De formule zet dat aantal om in een maandbedrag en past een binaire drempelwaarde toe die een beslissingsclassificatie afdwingt — geen narratieve aanbeveling die uitgesteld kan worden.

Let op de inherente beperking: het aantal is slechts een observatie van zichtbare vertrekken. Klanten die in de rij gaan staan en wachten (verslechterde ervaring, lagere terugkeerkans) worden niet meegerekend. Het verliescijfer is daarom een ondergrens, geen bovengrens.

<details>
<summary>Formule</summary>

```excel
Monthly_Revenue_Leakage = Lost_Customers_Daily × ASP × 30

Alert threshold:
=IF(Monthly_Revenue_Leakage > 500, "IMMEDIATE ACTION REQUIRED", "MONITOR")
```
</details>

---

## Voorbeeldscenario

**Locatie:** Onafhankelijke stedelijke espressobar, focus op afhaal in het stadscentrum.  
**Winkeltype:** Speed-First — Operaties wegen 40% mee, CX 30% en omzet 30%.  
**ASP:** $4,80 · **Waargenomen verloren klanten:** 6 per dag.

**Ruwe scores — ochtendpiek (alleen dienst 1 ter illustratie):**

| Metriek | Score | Anker uit de beoordelingsrubric |
|---|---|---|
| CX_001 Vriendelijkheid van de bediening | 4 | Warm en betrokken; persoonlijke aandacht waargenomen |
| CX_002 Wachtrij & wachttijd | 3 | Wachttijd 4–6 min; geen proactieve communicatie over de verwachte wachttijd |
| CX_003 Nauwkeurigheid van bestellingen | 5 | Geen fouten vastgesteld bij alle bestellingen |
| CX_004 Sfeer & zitplaatsen | 4 | Aangenaam; goede beschikbaarheid van zitplaatsen |
| OPS_001 Workflow van de barista | 4 | Efficiënt; minimale onnodige bewegingen |
| OPS_002 Betrouwbaarheid van apparatuur | 3 | Functioneel; kleine onderhoudsvertragingen |
| OPS_003 Netheid van de toonbank | 4 | Consistent schoon gedurende de hele dienst |
| OPS_004 Voorraadbeschikbaarheid | 3 | 1–2 artikelen niet beschikbaar; personeel informeerde klanten |
| REV_001 Zichtbaarheid van het menu | 3 | Menu zichtbaar; af en toe pogingen tot extra verkoop |
| REV_002 Piekcapaciteit | **2** | 40–60% bezet; geen promotiestrategie |
| REV_003 Prijsperceptie | 4 | Waarde duidelijk gecommuniceerd; premium-signalen aanwezig |
| REV_004 Secundaire bestedingen | 3 | Af en toe benoemt personeel aanvullende producten |

**Berekende uitkomsten:**

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

**Analytische interpretatie.** Cijfer B geeft aan dat de exploitatie over het geheel aan de basisnormen voldoet. Het maandelijkse omzetverlies van $864 en het waarschuwingssignaal bij REV_002 passen niet bij een uitgestelde actie. De onderbenutting van capaciteit (40–60% op piekmomenten) is de directe oorzaak van het verlies. De samengestelde score maskeert dit: op 3,50 ziet de locatie er stabiel uit; het via FILTER geëxtraheerde waarschuwingssignaal brengt het specifieke falenpunt naar boven.

**Aanbeveling.** Inzet van een eenvoudig wachtrij- of reserveringssysteem, gecombineerd met een promotie die in de spitsuren meer klanten trekt. Geen kapitaaluitgaven nodig.

**Implicatie voor de beslissing.** Als REV_002 na interventie van 2,0 naar 3,5 stijgt, stijgt het gemiddelde van de dimensie Omzetkans van 3,00 naar 3,38. De samengestelde score verschuift van 3,50 naar 3,61 — geen cijferwijziging, maar een betekenisvolle beweging richting de grens van categorie A (4,0) die in de volgende auditcyclus traceerbaar wordt.

**Onzekerheidsopmerking.** Deze projectie gaat ervan uit dat dienst 2 en 3 gelijke of betere scores opleveren. Een volledige audit over drie diensten kan uitwijzen dat de middagspits op Operaties slechter presteert dan de ochtendpiek, wat de samengestelde score zou dempen. Projecties op basis van één dienst moeten als indicatief worden gezien, niet als precies.

---

## Wat het werkboek werkelijk levert

- Een **uniforme vergelijkbare samengestelde score** die rekening houdt met het bedrijfsmodel, niet alleen met ruwe gemiddelden — vergelijkbaar tussen bezoeken en tussen locaties
- Een **prestatiecijfer** (S / A / B / C / F) dat de klasse in één teken communiceert, bruikbaar in klantpresentaties, bestuursdossiers en gesprekken over heronderhandeling van de huur
- Een **cijfer voor financieel verlies** dat observatie omzet in begrotingstaal en een binaire actiedrempel activeert bij $500 per maand
- Een **automatisch gegenereerde actielijst** die alleen falende metrieken bevat — geen handmatige filtering, geen punten gemist omdat ze tijdens de nabespreking niet opvielen
- Een **verplichting tot drie diensten** die vooroordeel door een enkelvoudig monster elimineert zonder de complexiteit van het bezoek te vergroten
- Een **systeem van gewichten per winkeltype** dat voorkomt dat een afhaalcafé wordt beoordeeld tegen de normen van een buurtcafé
- Een **beoordelingsrubric met 60 items** die twee onafhankelijke auditors die hetzelfde café scoren op binnen ongeveer 0,5 punt met elkaar in lijn brengt

---

## Werkboeklogica (voor de sceptische beoordelaar)

**Bladarchitectuur — strikt eenrichtingsdata-stroom, geen circulaire verwijzingen, geen macro's:**

```
Setup_Parameters  →  Input_Data  →  Calc_Engine [hidden]  →  Dashboard_Report
     (config)          (field obs)       (processing)            (output)
```

**Setup_Parameters** bevat: een gewichtsmatrix per winkeltype (Speed-First versus Space-First over 3 dimensies), een beoordelingsrubric met 60 items met de sleutel `Metric_ID_Score` (bijv. `CX_001_3`), benchmarkdoelstellingen voor de sector (CX: 4,0, Ops: 3,8, Rev: 3,5) en ASP-configuratie.

**Input_Data** verzamelt: 12 metrieekscores × 3 diensten (36 gescoorde cellen), rubricbeschrijvingen via XLOOKUP die automatisch verschijnen zodra een score wordt ingevoerd, en één financiële invoer (waargenomen verloren klanten per dag). Alle scorecellen worden via gegevensvalidatie met keuzelijsten beperkt tot gehele getallen van 1–5, met inline tooltips uit de beoordelingsrubric.

**Calc_Engine** (verborgen bij oplevering aan de klant) voert de volledige verwerkingsketen uit: dimensiegemiddelden per dienst → AVERAGEIF voor de totale gemiddelden → SUMPRODUCT voor de gewogen samengestelde score → IFS voor de cijferclassificatie → vermenigvuldiging voor het financieel verlies → booleaanse vector met waarschuwingssignalen (TRUE waar het gemiddelde < 2,5 ligt). De uitkomsten worden via benoemde bereiken door het dashboard gebruikt.

**Dashboard_Report** toont: drie KPI-kaarten (samengestelde score, cijfer, maandverlies met waarschuwingsdrempel van $500) → dimensiescore-pillen met verschil ten opzichte van de benchmark → een actiepaneel op basis van FILTER dat alleen metrieken met een waarschuwingssignaal toont, met kant-en-klare oplossingen → een plaatshouder voor een radardiagram die werkelijke waarden tegen de benchmark afzet over drie assen. Het afdrukbereik staat vast op A1:H45, A4 staand, passend op de pagina.

**Geregistreerde benoemde bereiken (14):** `Store_Type`, `ASP_Value`, `Speed_Weights`, `Space_Weights`, `Benchmark_Scores`, `Overall_Score`, `Performance_Grade`, `Monthly_Leakage`, `Metric_Report_Range`, `Red_Flag_Vector`, `CX_Score_Overall`, `Ops_Score_Overall`, `Rev_Score_Overall`, `Lost_Customers_Daily`.

---

## Implementatieopmerkingen

<details>
<summary>Formule voor de gewogen samengestelde score</summary>

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
<summary>Prestatiecijfer</summary>

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
<summary>Extractie van waarschuwingssignalen — FILTER-formule van het dashboard</summary>

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
<summary>XLOOKUP-rubricbeschrijving — live criteriumweergave in Input_Data</summary>

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
<summary>Radardiagram — handmatige opbouw (Bucket B)</summary>

Het werkboek bevat een grijze plaatshouder in rijen 25–44 van het dashboard met volledige bouwinstructies. Het diagram kan niet programmatisch worden gegenereerd via openpyxl.

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
<summary>Validatieregels en opleveringschecklist</summary>

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

## Het werkboek verkrijgen

Download `cafe_audit_tool.xlsx` uit deze repository.

Open in Excel voor desktop, iOS, Android of Online. Geen macro's. Geen setup nodig behalve het invullen van Winkeltype (Setup_Parameters B4) en ASP (Setup_Parameters E4).

Het simulatiebestand `cafe_audit_demo.html` spiegelt alle Calc_Engine-logica in de browser. Open het lokaal om de formule-uitkomsten te controleren tegen het Excel-werkboek voordat u het aan de klant oplevert.

---

## Beperkingen

**Wat dit werkboek niet kan:**

- Het houdt geen door klanten gerapporteerde ervaring vast. Scores weerspiegelen alleen de observatie van de auditor. Een café kan 4,5 scoren op CX terwijl klanten 2,8 op NPS aangeven — observatie en perceptie zijn geen gelijkwaardige metrieken.
- Het houdt geen longitudinale gegevens bij. Elke audit is een momentopname. Trendanalyse vereist het handmatig vastleggen van samengestelde scores en verliescijfers over bezoeken heen in een aparte tracker.
- De berekening van het financieel verlies gaat ervan uit dat alle verloren klanten hetzelfde ASP hebben. Als vertrekkende klanten juist veel besteden, is de schatting conservatief. Als ze juist weinig besteden (bijv. kijkers in plaats van kopers), wordt het verlies overschat.
- De benchmarkwaarden (CX: 4,0, Ops: 3,8, Rev: 3,5) en de waarschuwingsdrempel van $500 zijn instelbare standaardwaarden, geen statistisch gevalideerde sectorgemiddelden. Gebruikers in stedelijke markten met hoge volumes of voorstedelijke contexten met lage volumes moeten Setup_Parameters!B14:B16 en de verliesdrempel zelf opnieuw ijken.
- De drempelwaarde van 2,5 voor waarschuwingssignalen is een snede in het midden van de schaal. Het brengt metrieken die 2,6 scoren niet naar boven, terwijl die in context nog steeds operationeel risico kunnen vertegenwoordigen.

---

## Andere beslissingstools die u nodig zou kunnen hebben

| Scenario | Voorgesteld instrument |
|---|---|
| Vergelijking van een portefeuille met meerdere locaties (10+ café's) | Verzamel auditscores in een aparte Excel-tracker met GROUPBY/PIVOTBY over de kolommen samengestelde score en verlies |
| Analyse van het gat tussen klantperceptie en operationele score | Combineer de samengestelde auditscore met NPS- of CSAT-gegevens; correlatieanalyse in een apart model |
| Meting van het rendement van personeelstraining | Pre/post-audit met dit werkboek; het verschil in samengestelde score en maandverlies is de belangrijkste KPI |
| Ondersteuning bij een beslissing over huurverlenging | Combineer de samengestelde score met omzet per vierkante meter en bezoekersaantallen in een apart financieel model |
| Analyse van menuprijzen | Vereist transactiedata op POS-niveau; valt buiten de scope van een op observatie gebaseerd auditinstrument |
| Benchmarking tegen genoemde merken | Vereist externe benchmarkgegevens; de benchmarks in het werkboek zijn instelbaar maar worden niet automatisch bijgewerkt |

---

## Licentie

Copyright 2025. Uitgebracht onder de [Apache License 2.0](https://www.apache.org/licenses/LICENSE-2.0).

Gebruik, wijzig en distribueer dit werkboek en bijbehorende bestanden onder de voorwaarden van de Apache License 2.0. Naamsvermelding vereist. Er wordt geen garantie geboden, expliciet noch impliciet. Zie het bestand `LICENSE` voor de volledige voorwaarden.

---

*Beslissingondersteunend instrument. Geen bedrijfssysteem.*

---

## Café-audit · 3D-diagnostische motor

[![Licentie](https://img.shields.io/badge/许可证-Apache_2.0-blue.svg)](https://opensource.org/licenses/Apache-2.0)
[![Platform](https://img.shields.io/badge/平台-Excel%20%7C%20Web-navy)](https://github.com)
[![Type tool](https://img.shields.io/badge/工具类型-决策支持-2251FF)](https://github.com)

**Een gestructureerd auditraamwerk dat observaties ter plaatse van een café omzet in een gewogen diagnostische score, een schatting van het omzetverlies en een geprioriteerde actielijst — in minder dan 30 minuten per bezoek.**

De inhoud van deze Nederlandstalige versie is consistent met de Engelse versie. Alle formules, gewogslogica, architectuurbeschrijvingen en beperkingsbepalingen zijn van toepassing. Volledige technische details staan in de uitvouwbare blokken van de Engelse versie.

[▶ onlinedemo](cafe_audit_demo.html) · [⬇ Download werkboek](cafe_audit_tool.xlsx)
