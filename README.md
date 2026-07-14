
# Café Audit · 3D Diagnostics Engine

[![License](https://img.shields.io/badge/License-Apache_2.0-blue.svg)](https://opensource.org/licenses/Apache-2.0)
[![Platform](https://img.shields.io/badge/Platform-Excel%20%7C%20Web-navy)](https://github.com)
[![Tool Type](https://img.shields.io/badge/Tool_Type-Decision_Support-2251FF)](https://github.com)

**A structured audit framework that converts on-site café observation into a weighted diagnostic score, financial leakage estimate, and prioritised action list — in under 30 minutes per visit.**

---

[▶ Live Demo](https://hyvoid.github.io/Caf-audit-tool-for-efficient-reviews/) · [⬇ Download Workbook](cafe_audit_tool.xlsx) · 

---

 <img width="1536" height="1024" alt="ChatGPT Image Jul 14, 2026, 06_20_33 PM" src="https://github.com/user-attachments/assets/99030acd-62a9-4f24-b7dd-425c643c8a39" />


---

## What Decision Does This Help You Make?

1. **Which operational dimension is costing the most revenue** — Customer Experience, Operations, or Revenue Opportunity — weighted by store type?
2. **Is this café performing above or below the industry minimum threshold** — and by how much, expressed as a composite score on a verified 5-point scale?
3. **How much revenue is leaking monthly** from observable failures, and does that number cross the $500 intervention threshold?
4. **Which specific metrics require immediate action** and what is the first corrective step for each?

---

## About The Builder

Three operating constraints govern every design decision in this workbook.

**Decision support over system replacement.** This is not a POS integration, a loyalty analytics platform, or an operations management suite. It is a structured observation instrument that produces a decision-ready output from field data collected on a tablet in a single visit.

**Productized reasoning over feature accumulation.** Every formula serves a decision. The SUMPRODUCT weighting engine exists because unweighted averages systematically misrepresent performance for cafés with different operating models. The FILTER-based action list exists because a dashboard without prioritisation still requires human judgment to act.

**Low-friction execution over implementation complexity.** No macros. No VBA dependencies. No external data connections. The workbook runs on Excel for iOS, Android, iPadOS, and Excel Online without modification.

---

## Who This Is For

| User | Typical Situation | Decision Supported | Why Existing Methods Fail |
|---|---|---|---|
| Independent café consultant | Auditing 4–12 sites per month for a portfolio owner | Which sites need capital, which need process, which need staff | Gut-feel visits produce inconsistent outputs; clients cannot compare across visits |
| F&B operations manager | Comparing branch performance for quarterly review | Where to focus regional manager attention | Branch-level sales data does not isolate root cause; two cafés with identical revenue can have opposite operational health |
| Café owner / operator | Self-diagnosing before a lease renewal or refit decision | Whether the problem is product, people, or space | Owner proximity creates blind spots; no external reference point for scores |
| Hospitality training provider | Pre/post training assessment for client cafés | Whether training produced measurable operational uplift | Pre/post comparison requires a consistent measurement instrument, which most operators do not have |

---

## Why Most Café Audit Errors Aren't Judgment Errors

The auditor usually has sufficient expertise. The error enters earlier — in the measurement structure.

**Where the process breaks:**

```
Unstructured observation
    → inconsistent metric coverage per visit
    → no rubric → scores reflect auditor mood, not café performance
    → averages across dimensions hide the worst-performing metric
    → action list ranked by salience, not by weighted impact
    → owner receives a report; no decision follows
```

**With a structured instrument:**

```
Rubric-anchored 5-point scoring (12 metrics × 3 shifts)
    → AVERAGE per metric across shifts → removes single-event bias
    → SUMPRODUCT × store-type weights → composite reflects operating model
    → IFS grade → single comparable output
    → FILTER(Red_Flag_Vector = TRUE) → action list auto-extracts
    → Lost customers × ASP × 30 → converts observation to financial stakes
```

**Operational implication:** An auditor who scores "Queue & Wait Time" as 3 on intuition may score it 2 when forced to answer: *"Is the wait 4–6 minutes with no proactive ETA communication?"* The rubric removes the gap between two auditors observing the same queue.

The workbook does not make the judgment. It prevents judgment from being made without a definition.

---

## Three Traps That Catch Even Experienced Operators

### Trap 1 — The Unweighted Average Misrepresents the Operating Model

**A decision was made.** A café scores 3.8 on Customer Experience and 2.9 on Operations. The unweighted average across three dimensions is 3.30. The operator concludes performance is acceptable and defers intervention.

**The unnoticed flaw.** The café is a takeaway-focused Speed-First operation. Operations carries a 40% weight under Speed-First, not the implied 33.3% of an unweighted average.

| Method | CX (3.8) | Ops (2.9) | Rev (3.2) | Composite | Grade |
|---|---|---|---|---|---|
| Unweighted average | 33.3% | 33.3% | 33.3% | **3.30** | B |
| Speed-First weights | 30% | **40%** | 30% | **3.20** | B |
| Space-First weights | **50%** | 30% | 20% | **3.42** | B |

The grade holds here — but at the 3.0 boundary (B/C), a 0.10 shift is the difference between a deferred decision and an immediate operational review. More importantly, the weight difference signals *where* to intervene: Speed-First operators must treat a 2.9 Operations score as the primary risk, not a secondary concern balanced against CX.

**Corrected approach.** Apply store-type weights before comparing composites across sites. A Space-First café with identical raw scores sits at a different risk tier because CX carries 20 additional percentage points.

<details>
<summary>Formula</summary>

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

### Trap 2 — The Single-Shift Sample Inflates the Score

**A decision was made.** An auditor visits during Morning Peak, observes a 4.5 on Barista Workflow Efficiency, and reports strong operations. The owner decides no workflow intervention is needed.

**The unnoticed flaw.** Morning Peak is the highest-preparation shift. Off-Peak staffing is leaner. The same metric scores 2.0 at 18:30 because one barista covers both counter and drinks with no defined handoff protocol.

**Before — single-shift audit:**

```
OPS_001 Morning Peak only:  4.5
Reported average:           4.5
Decision:                   No action required
```

**After — three-shift audit:**

```
OPS_001 Morning Peak:    4.5
OPS_001 Afternoon Rush:  3.5
OPS_001 Off-Peak:        2.0
Three-shift average:     3.33
Red flag:                NOT triggered (3.33 ≥ 2.5 threshold)
Additional finding:      Off-Peak requires a targeted staffing review
```

**Why the reasoning is incorrect.** Single-shift audits sample the best-prepared window by default. Morning Peak staffing levels are typically higher and pre-briefed. The instrument mandates three operational windows precisely to prevent this selection bias from entering the composite.

**Corrected outcome.** The AVERAGEIF formula excludes unscored shifts using a `>0` condition. No missing shift data distorts the composite toward zero — but the Off-Peak score alone is a flaggable data point for the recommendation section even when the average clears the threshold.

<details>
<summary>Formula</summary>

```excel
-- Per-metric, per-shift average across the three shift columns:
=IFERROR(AVERAGEIF(B_row:D_row, ">0"), "")

-- Excludes shifts where no score was entered (value = 0)
-- A shift left unscored does not count as a zero in the average
```
</details>

---

### Trap 3 — Revenue Leakage Stays Qualitative

**A decision was made.** An auditor notes "some customers left due to queue length" in a narrative report. The owner acknowledges the observation and takes no action because the dollar impact is unclear.

**The unnoticed flaw.** "Some customers" is not actionable. At 8 observed departures per day with an average spend of $5.50, the monthly leakage is $1,320 — 2.6 times the $500 intervention threshold. The observation warranted immediate queue management deployment, not a narrative footnote deferred to next quarter.

**Before — qualitative:**

```
Finding:        Queue management below standard
Evidence:       "Customers leaving during peak"
Recommendation: "Consider reviewing queue process"
Owner response: Noted for Q3 planning
```

**After — quantified:**

```
Lost customers (observed, daily):  8
ASP:                               $5.50
Monthly leakage:                   8 × $5.50 × 30 = $1,320
Alert threshold exceeded:          YES (> $500)
Required decision:                 Queue intervention; ROI positive within ~2 weeks
```

**Corrected approach.** The auditor counts observable walk-aways during the visit. The formula converts that count to a monthly figure and applies a binary threshold that forces a decision classification — not a narrative recommendation that can be deferred.

Note the inherent limitation: the count is an observation of visible departures only. Customers who join and wait (degraded experience, reduced return probability) are not captured. The leakage figure is therefore a floor estimate, not a ceiling.

<details>
<summary>Formula</summary>

```excel
Monthly_Revenue_Leakage = Lost_Customers_Daily × ASP × 30

Alert threshold:
=IF(Monthly_Revenue_Leakage > 500, "IMMEDIATE ACTION REQUIRED", "MONITOR")
```
</details>

---

## Example Scenario

**Site:** Independent urban espresso bar, city-centre takeaway focus.  
**Store type:** Speed-First — Operations weighted 40%, CX 30%, Revenue 30%.  
**ASP:** $4.80 · **Observed lost customers:** 6 per day.

**Raw scores — Morning Peak (Shift 1 only for illustration):**

| Metric | Score | Rubric Anchor |
|---|---|---|
| CX_001 Service Friendliness | 4 | Warm and engaged; personal touch observed |
| CX_002 Queue & Wait Time | 3 | Wait 4–6 min; no proactive ETA communication |
| CX_003 Order Accuracy | 5 | Zero errors confirmed across all orders |
| CX_004 Atmosphere & Seating | 4 | Pleasant; good seating availability |
| OPS_001 Barista Workflow | 4 | Efficient; minimal waste motion |
| OPS_002 Equipment Reliability | 3 | Functional; minor maintenance delays |
| OPS_003 Counter Cleanliness | 4 | Consistently clean throughout shift |
| OPS_004 Stock Availability | 3 | 1–2 items unavailable; staff informed customers |
| REV_001 Menu Visibility | 3 | Menu visible; occasional upsell attempts |
| REV_002 Peak Capacity | **2** | 40–60% occupied; no promotional strategy |
| REV_003 Pricing Perception | 4 | Value clearly communicated; premium signals present |
| REV_004 Secondary Spend | 3 | Occasional staff mention of add-ons |

**Calculated outputs:**

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

**Analytical interpretation.** Grade B indicates the operation meets baseline standards overall. The $864 monthly leakage and REV_002 red flag are inconsistent with a deferred-action posture. The capacity underutilisation (40–60% at peak) is the proximate cause of the leakage. The composite score masks this: at 3.50, the site looks stable; the FILTER-extracted red flag surfaces the specific failure point.

**Recommendation.** Queue-ahead or light reservation system deployment, combined with a peak-hour footfall promotion. No capex required.

**Decision implication.** If REV_002 improves from 2.0 to 3.5 following intervention, Revenue dimension average rises from 3.00 to 3.38. Composite shifts from 3.50 to 3.61 — not a grade change, but a meaningful movement toward the A-tier boundary (4.0) that becomes trackable in the next audit cycle.

**Uncertainty note.** This projection assumes Shifts 2 and 3 produce equivalent or better scores. A full three-shift audit may reveal that Afternoon Rush underperforms Morning Peak on Operations, which would dampen the composite. Single-shift projections should be treated as directional, not precise.

---

## What The Workbook Actually Delivers

- A **single comparable composite score** that accounts for operating model, not just raw averages — comparable across visits and across sites
- A **performance grade** (S / A / B / C / F) that communicates tier in one character, usable in client decks, board packs, and lease renegotiation conversations
- A **financial leakage figure** that converts observation into budget language and triggers a binary action threshold at $500/month
- An **auto-generated action list** that extracts only failing metrics — no manual filtering, no items missed because they weren't salient during the debrief
- A **three-shift coverage mandate** that eliminates single-sample bias without increasing visit complexity
- A **store-type weight system** that prevents a takeaway bar from being evaluated against community café standards
- A **60-entry rubric dictionary** that aligns two independent auditors scoring the same café to within approximately 0.5 points

---

## Workbook Logic (For the Skeptical Reviewer)

**Sheet architecture — strict single-direction data flow, no circular references, no macros:**

```
Setup_Parameters  →  Input_Data  →  Calc_Engine [hidden]  →  Dashboard_Report
     (config)          (field obs)       (processing)            (output)
```

**Setup_Parameters** contains: store-type weight matrix (Speed-First vs Space-First across 3 dimensions), 60-entry rubric dictionary keyed as `Metric_ID_Score` (e.g. `CX_001_3`), industry benchmark targets (CX: 4.0, Ops: 3.8, Rev: 3.5), and ASP configuration.

**Input_Data** collects: 12 metric scores × 3 shifts (36 scored cells), XLOOKUP-driven rubric descriptions that auto-populate on score entry, and one financial input (observed lost customers per day). All score cells are constrained by data validation dropdowns to integers 1–5 with inline rubric tooltips.

**Calc_Engine** (hidden in client delivery) executes the full processing chain: per-shift dimension averages → AVERAGEIF overall averages → SUMPRODUCT weighted composite → IFS grade classification → financial leakage multiplication → boolean red-flag vector (TRUE where avg < 2.5). Outputs are consumed by Dashboard via named ranges.

**Dashboard_Report** renders: three KPI cards (composite score, grade, monthly leakage with $500 alert threshold) → dimension score pills with benchmark delta → FILTER-powered action panel listing only red-flag metrics with pre-written remedies → radar chart placeholder comparing actual vs benchmark across three axes. Print area fixed to A1:H45, A4 portrait, fit-to-page.

**Named ranges registered (14):** `Store_Type`, `ASP_Value`, `Speed_Weights`, `Space_Weights`, `Benchmark_Scores`, `Overall_Score`, `Performance_Grade`, `Monthly_Leakage`, `Metric_Report_Range`, `Red_Flag_Vector`, `CX_Score_Overall`, `Ops_Score_Overall`, `Rev_Score_Overall`, `Lost_Customers_Daily`.

---

## Implementation Notes

<details>
<summary>Weighted composite formula</summary>

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
<summary>Performance grade</summary>

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
<summary>Red flag extraction — Dashboard FILTER formula</summary>

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
<summary>XLOOKUP rubric description — Input_Data live criteria display</summary>

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
<summary>Radar chart — Bucket B manual build</summary>

The workbook contains a grey placeholder in Dashboard rows 25–44 with full build instructions. The chart cannot be generated programmatically via openpyxl.

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
<summary>Validation rules and delivery checklist</summary>

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

## Get The Workbook

Download `cafe_audit_tool.xlsx` from this repository.

Open in Excel desktop, iOS, Android, or Online. No macros. No setup required beyond entering Store Type (Setup_Parameters B4) and ASP (Setup_Parameters E4).

The simulation file `cafe_audit_demo.html` mirrors all Calc_Engine logic in-browser. Open it locally to validate formula outputs against the Excel workbook before client delivery.

---

## Limitations

**What this workbook cannot do:**

- It does not capture customer-reported experience. Scores reflect auditor observation only. A café can score 4.5 on CX while customers report 2.8 on NPS — observation and perception are not equivalent metrics.
- It does not track longitudinal data. Each audit is a single-point snapshot. Trend analysis requires manual logging of composite scores and leakage figures across visits in a separate tracker.
- The financial leakage calculation assumes all lost customers carry equal ASP. If departing customers skew high-spend, the estimate is conservative. If they skew low-spend (e.g., browsers rather than buyers), it overstates leakage.
- The benchmark values (CX: 4.0, Ops: 3.8, Rev: 3.5) and the $500 alert threshold are configurable defaults, not statistically validated industry averages. Users in high-volume urban markets or low-volume suburban contexts must recalibrate Setup_Parameters!B14:B16 and the leakage threshold independently.
- The red-flag threshold of 2.5 is a mid-scale cut. It will not surface metrics scoring 2.6 that may still represent operational risk in context.

---

## Other Decision Tools You Might Need

| Scenario | Suggested Instrument |
|---|---|
| Multi-site portfolio comparison (10+ cafés) | Aggregate audit scores into a separate Excel tracker using GROUPBY/PIVOTBY across the composite score and leakage columns |
| Customer perception vs operational score gap analysis | Pair audit composite with NPS or CSAT data; correlation analysis in a separate model |
| Staff training ROI measurement | Pre/post audit using this workbook; delta in composite score and monthly leakage is the primary KPI |
| Lease renewal decision support | Combine composite score with revenue per square metre and footfall count in a separate financial model |
| Menu pricing analysis | Requires transaction-level POS data; outside the scope of an observation-based audit instrument |
| Competitive benchmarking against named brands | Requires externally sourced benchmark data; the workbook benchmarks are configurable but not auto-updated |

---

## License

Copyright 2025. Released under the [Apache License 2.0](https://www.apache.org/licenses/LICENSE-2.0).

Use, modify, and distribute this workbook and associated files under the terms of the Apache License 2.0. Attribution required. No warranty is provided, express or implied. See the `LICENSE` file for full terms.

---

*Decision support instrument. Not an enterprise system.*

---

# 咖啡馆审计 · 三维诊断引擎

[![许可证](https://img.shields.io/badge/许可证-Apache_2.0-blue.svg)](https://opensource.org/licenses/Apache-2.0)
[![平台](https://img.shields.io/badge/平台-Excel%20%7C%20Web-navy)](https://github.com)
[![工具类型](https://img.shields.io/badge/工具类型-决策支持-2251FF)](https://github.com)

**一款结构化审计框架，将现场咖啡馆观察数据转化为加权诊断评分、财务漏损估算和优先行动清单——单次审计在 30 分钟内完成。**

本中文版内容与英文版保持一致。所有公式、权重逻辑、架构说明和局限性条款均适用。完整技术细节见英文版各折叠区块。

[▶ 在线演示](cafe_audit_demo.html) · [⬇ 下载工作簿](cafe_audit_tool.xlsx)
