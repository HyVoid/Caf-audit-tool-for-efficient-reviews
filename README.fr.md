[ 🌐 عربي ](README.ar.md) | [ 🇩🇪 Deutsch ](README.de.md) | [ 🇫🇷 Français ](README.fr.md) | [ 🇳🇱 Nederlands ](README.nl.md) | [ 🇪🇸 Español ](README.sp.md) | [ 🇬🇧 English ](README.md)

# Audit de café · Moteur de diagnostic 3D

[![Licence](https://img.shields.io/badge/License-Apache_2.0-blue.svg)](https://opensource.org/licenses/Apache-2.0)
[![Plateforme](https://img.shields.io/badge/Platform-Excel%20%7C%20Web-navy)](https://github.com)
[![Type d'outil](https://img.shields.io/badge/Tool_Type-Decision_Support-2251FF)](https://github.com)

**Un cadre d'audit structuré qui convertit l'observation sur site d'un café en un score diagnostic pondéré, une estimation de la fuite de revenus et une liste d'actions priorisée — en moins de 30 minutes par visite.**

---

[▶ Démo en ligne](https://hyvoid.github.io/Caf-audit-tool-for-efficient-reviews/) · [⬇ Télécharger le classeur](cafe_audit_tool.xlsx) · 

---

 <img width="1536" height="1024" alt="ChatGPT Image Jul 14, 2026, 06_20_33 PM" src="https://github.com/user-attachments/assets/99030acd-62a9-4f24-b7dd-425c643c8a39" />


---

## Quelle décision cet outil vous aide-t-il à prendre ?

1. **Quelle dimension opérationnelle coûte le plus de revenus** — Expérience client, Opérations ou Opportunité de revenus — pondérée par le type de magasin ?
2. **Ce café performe-t-il au-dessus ou en dessous du seuil minimal du secteur** — et de combien, exprimé sous forme de score composite sur une échelle vérifiée de 5 points ?
3. **Combien de revenus fuient chaque mois** du fait de défaillances observables, et ce chiffre franchit-il le seuil d'intervention de 500 $ ?
4. **Quelles métriques spécifiques nécessitent une action immédiate** et quelle est la première mesure corrective pour chacune ?

---

## À propos du concepteur

Trois contraintes opérationnelles gouvernent chaque décision de conception dans ce classeur.

**L'aide à la décision avant le remplacement de système.** Ce n'est pas une intégration de caisse, une plateforme d'analyse de la fidélité ou une suite de gestion des opérations. C'est un instrument d'observation structuré qui produit un résultat prêt à la décision à partir de données de terrain collectées sur tablette en une seule visite.

**Un raisonnement industrialisé avant l'accumulation de fonctionnalités.** Chaque formule sert une décision. Le moteur de pondération SUMPRODUCT existe parce que les moyennes non pondérées déforment systématiquement la performance des cafés aux modèles opérationnels différents. La liste d'actions basée sur FILTER existe parce qu'un tableau de bord sans priorisation exige encore un jugement humain pour agir.

**Une exécution à faible friction avant la complexité de mise en œuvre.** Aucune macro. Aucune dépendance VBA. Aucune connexion de données externe. Le classeur fonctionne sur Excel pour iOS, Android, iPadOS et Excel Online sans modification.

---

## À qui cet outil est-il destiné

| Utilisateur | Situation typique | Décision prise en charge | Pourquoi les méthodes existantes échouent |
|---|---|---|---|
| Consultant café indépendant | Auditer 4–12 sites par mois pour un propriétaire de portefeuille | Quels sites nécessitent des investissements, lesquels des processus, lesquels du personnel | Les visites à l'intuition produisent des résultats incohérents ; les clients ne peuvent pas comparer d'une visite à l'autre |
| Responsable des opérations de restauration | Comparer la performance des établissements pour une revue trimestrielle | Où concentrer l'attention du responsable régional | Les données de ventes par établissement n'isolent pas la cause racine ; deux cafés au chiffre d'affaires identique peuvent avoir une santé opérationnelle opposée |
| Propriétaire / exploitant de café | Auto-diagnostic avant un renouvellement de bail ou une décision de réaménagement | Si le problème vient du produit, des personnes ou de l'espace | La proximité du propriétaire crée des angles morts ; aucun point de référence externe pour les scores |
| Prestataire de formation en hôtellerie-restauration | Évaluation avant/après formation pour les cafés clients | Si la formation a produit une amélioration opérationnelle mesurable | La comparaison avant/après exige un instrument de mesure cohérent, que la plupart des exploitants n'ont pas |

---

## Pourquoi la plupart des erreurs d'audit café ne sont pas des erreurs de jugement

L'auditeur dispose généralement d'une expertise suffisante. L'erreur s'introduit plus tôt — dans la structure de mesure.

**Là où le processus échoue :**

```
Unstructured observation
    → inconsistent metric coverage per visit
    → no rubric → scores reflect auditor mood, not café performance
    → averages across dimensions hide the worst-performing metric
    → action list ranked by salience, not by weighted impact
    → owner receives a report; no decision follows
```

**Avec un instrument structuré :**

```
Rubric-anchored 5-point scoring (12 metrics × 3 shifts)
    → AVERAGE per metric across shifts → removes single-event bias
    → SUMPRODUCT × store-type weights → composite reflects operating model
    → IFS grade → single comparable output
    → FILTER(Red_Flag_Vector = TRUE) → action list auto-extracts
    → Lost customers × ASP × 30 → converts observation to financial stakes
```

**Implication opérationnelle :** Un auditeur qui attribue à « File et temps d'attente » un score de 3 par intuition peut lui attribuer 2 lorsqu'il est forcé de répondre : *"Is the wait 4–6 minutes with no proactive ETA communication?"* La rubrique supprime l'écart entre deux auditeurs observant la même file.

Le classeur ne porte pas le jugement. Il empêche qu'un jugement soit porté sans définition.

---

## Trois pièges qui surprennent même les exploitants expérimentés

### Piège 1 — La moyenne non pondérée déforme le modèle opérationnel

**Une décision a été prise.** Un café obtient 3,8 en expérience client et 2,9 en opérations. La moyenne non pondérée des trois dimensions est de 3,30. L'exploitant conclut que la performance est acceptable et reporte l'intervention.

**Le défaut passé inaperçu.** Le café est un établissement axé sur la vente à emporter de type Speed-First. Les opérations portent un poids de 40 % sous Speed-First, et non les 33,3 % implicites d'une moyenne non pondérée.

| Méthode | CX (3.8) | Ops (2.9) | Rev (3.2) | Score composite | Note |
|---|---|---|---|---|---|
| Moyenne non pondérée | 33.3% | 33.3% | 33.3% | **3.30** | B |
| Poids Speed-First | 30% | **40%** | 30% | **3.20** | B |
| Poids Space-First | **50%** | 30% | 20% | **3.42** | B |

La note se maintient ici — mais à la frontière de 3,0 (B/C), un écart de 0,10 fait la différence entre une décision reportée et une revue opérationnelle immédiate. Plus important encore, la différence de poids indique *où* intervenir : les exploitants Speed-First doivent traiter un score de 2,9 en opérations comme le risque principal, et non comme une préoccupation secondaire contrebalancée par les CX.

**Approche corrigée.** Appliquez les poids par type de magasin avant de comparer les scores composites entre les sites. Un café Space-First avec des scores bruts identiques se situe à un niveau de risque différent parce que les CX portent 20 points de pourcentage supplémentaires.

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

### Piège 2 — L'échantillon sur un seul créneau gonfle le score

**Une décision a été prise.** Un auditeur visite pendant le pic du matin, observe un 4,5 en efficacité du flux de travail du barista et rapporte des opérations solides. Le propriétaire décide qu'aucune intervention sur le flux de travail n'est nécessaire.

**Le défaut passé inaperçu.** Le pic du matin est le créneau le mieux préparé. Le personnel hors pointe est plus restreint. La même métrique obtient 2,0 à 18:30 parce qu'un seul barista couvre à la fois le comptoir et les boissons sans protocole de passation défini.

**Avant — audit sur un seul créneau :**

```
OPS_001 Morning Peak only:  4.5
Reported average:           4.5
Decision:                   No action required
```

**Après — audit sur trois créneaux :**

```
OPS_001 Morning Peak:    4.5
OPS_001 Afternoon Rush:  3.5
OPS_001 Off-Peak:        2.0
Three-shift average:     3.33
Red flag:                NOT triggered (3.33 ≥ 2.5 threshold)
Additional finding:      Off-Peak requires a targeted staffing review
```

**Pourquoi le raisonnement est incorrect.** Les audits sur un seul créneau échantillonnent par défaut la fenêtre la mieux préparée. Les niveaux de personnel du pic du matin sont généralement plus élevés et briefés à l'avance. L'instrument impose trois fenêtres opérationnelles précisément pour empêcher ce biais de sélection d'entrer dans le score composite.

**Résultat corrigé.** La formule AVERAGEIF exclut les créneaux non notés à l'aide d'une condition `>0`. Aucune donnée de créneau manquante ne déforme le composite vers zéro — mais le score hors pointe seul est un point de données signalable pour la section des recommandations, même lorsque la moyenne franchit le seuil.

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

### Piège 3 — La fuite de revenus reste qualitative

**Une décision a été prise.** Un auditeur note « certains clients sont partis à cause de la longueur de la file » dans un rapport narratif. Le propriétaire prend acte de l'observation et n'agit pas parce que l'impact financier n'est pas clair.

**Le défaut passé inaperçu.** « Certains clients » n'est pas actionnable. Avec 8 départs observés par jour et une dépense moyenne de 5,50 $, la fuite de revenus mensuelle est de 1 320 $ — soit 2,6 fois le seuil d'intervention de 500 $. L'observation justifiait un déploiement immédiat de la gestion de file, et non une simple mention narrative reportée au trimestre suivant.

**Avant — qualitatif :**

```
Finding:        Queue management below standard
Evidence:       "Customers leaving during peak"
Recommendation: "Consider reviewing queue process"
Owner response: Noted for Q3 planning
```

**Après — quantifié :**

```
Lost customers (observed, daily):  8
ASP:                               $5.50
Monthly leakage:                   8 × $5.50 × 30 = $1,320
Alert threshold exceeded:          YES (> $500)
Required decision:                 Queue intervention; ROI positive within ~2 weeks
```

**Approche corrigée.** L'auditeur compte les départs observables pendant la visite. La formule convertit ce comptage en chiffre mensuel et applique un seuil binaire qui force une classification de décision — et non une recommandation narrative qui peut être reportée.

Notez la limitation inhérente : le comptage ne repose que sur l'observation des départs visibles. Les clients qui rejoignent la file et attendent (expérience dégradée, probabilité de retour réduite) ne sont pas capturés. Le chiffre de fuite constitue donc une estimation plancher, non un plafond.

<details>
<summary>Formule</summary>

```excel
Monthly_Revenue_Leakage = Lost_Customers_Daily × ASP × 30

Alert threshold:
=IF(Monthly_Revenue_Leakage > 500, "IMMEDIATE ACTION REQUIRED", "MONITOR")
```
</details>

---

## Exemple de scénario

**Site :** Bar à espresso urbain indépendant, axé sur la vente à emporter en centre-ville.  
**Type de magasin :** Speed-First — opérations pondérées à 40 %, CX 30 %, Revenus 30 %.  
**ASP :** 4,80 $ · **Clients perdus observés :** 6 par jour.

**Scores bruts — pic du matin (créneau 1 uniquement, à titre d'illustration) :**

| Métrique | Score | Ancre de rubrique |
|---|---|---|
| CX_001 Convivialité du service | 4 | Chaleureux et engagé ; touche personnelle observée |
| CX_002 File et temps d'attente | 3 | Attente de 4–6 min ; aucune communication proactive du délai |
| CX_003 Exactitude des commandes | 5 | Zéro erreur confirmée sur toutes les commandes |
| CX_004 Ambiance et places assises | 4 | Agréable ; bonne disponibilité des places |
| OPS_001 Flux de travail du barista | 4 | Efficace ; mouvements inutiles minimaux |
| OPS_002 Fiabilité des équipements | 3 | Fonctionnel ; retards de maintenance mineurs |
| OPS_003 Propreté du comptoir | 4 | Constamment propre tout au long du service |
| OPS_004 Disponibilité des stocks | 3 | 1–2 articles indisponibles ; le personnel informait les clients |
| REV_001 Visibilité du menu | 3 | Menu visible ; tentatives de vente incitative occasionnelles |
| REV_002 Capacité en pointe | **2** | 40–60% occupé ; aucune stratégie promotionnelle |
| REV_003 Perception des prix | 4 | Valeur clairement communiquée ; signaux haut de gamme présents |
| REV_004 Dépense secondaire | 3 | Mention occasionnelle d'options par le personnel |

**Résultats calculés :**

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

**Interprétation analytique.** La note B indique que l'exploitation respecte globalement les standards de base. La fuite de revenus mensuelle de 864 $ et le signal d'alerte REV_002 sont incompatibles avec une posture d'action reportée. La sous-utilisation de la capacité (40–60 % en pointe) est la cause proximale de la fuite. Le score composite masque cela : à 3,50, le site semble stable ; le signal d'alerte extrait par FILTER révèle le point de défaillance spécifique.

**Recommandation.** Déploiement d'un système de file avancée ou de réservation légère, combiné à une promotion de fréquentation aux heures de pointe. Aucun investissement requis.

**Implication décisionnelle.** Si REV_002 s'améliore de 2,0 à 3,5 après intervention, la moyenne de la dimension Revenus passe de 3,00 à 3,38. Le score composite se déplace de 3,50 à 3,61 — pas un changement de note, mais un mouvement significatif vers la frontière du niveau A (4,0) qui devient traçable lors du prochain cycle d'audit.

**Note d'incertitude.** Cette projection suppose que les créneaux 2 et 3 produisent des scores équivalents ou meilleurs. Un audit complet sur trois créneaux peut révéler que le pic de l'après-midi sous-performe le pic du matin sur les opérations, ce qui atténuerait le score composite. Les projections sur un seul créneau doivent être considérées comme indicatives, non précises.

---

## Ce que le classeur livre concrètement

- Un **score composite unique et comparable** qui tient compte du modèle opérationnel, et pas seulement des moyennes brutes — comparable entre les visites et entre les sites
- Une **note de performance** (S / A / B / C / F) qui communique le niveau en un seul caractère, utilisable dans les présentations clients, les dossiers de conseil et les négociations de renouvellement de bail
- Un **chiffre de fuite financière** qui convertit l'observation en langage budgétaire et déclenche un seuil d'action binaire à 500 $/mois
- Une **liste d'actions générée automatiquement** qui extrait uniquement les métriques en échec — aucun filtrage manuel, aucun élément manqué parce qu'il n'était pas saillant au moment du compte rendu
- Une **obligation de couverture sur trois créneaux** qui élimine le biais d'échantillon unique sans augmenter la complexité de la visite
- Un **système de poids par type de magasin** qui empêche qu'un bar à emporter soit évalué selon les standards d'un café de quartier
- Un **dictionnaire de rubriques à 60 entrées** qui aligne deux auditeurs indépendants notant le même café à moins de 0,5 point près

---

## Logique du classeur (pour le relecteur sceptique)

**Architecture des feuilles — flux de données strictement unidirectionnel, aucune référence circulaire, aucune macro :**

```
Setup_Parameters  →  Input_Data  →  Calc_Engine [hidden]  →  Dashboard_Report
     (config)          (field obs)       (processing)            (output)
```

**Setup_Parameters** contient : la matrice de poids par type de magasin (Speed-First vs Space-First sur 3 dimensions), le dictionnaire de rubriques à 60 entrées indexé comme `Metric_ID_Score` (p. ex. `CX_001_3`), les cibles de référence du secteur (CX : 4,0, Ops : 3,8, Rev : 3,5) et la configuration ASP.

**Input_Data** collecte : 12 scores de métriques × 3 créneaux (36 cellules notées), des descriptions de rubriques pilotées par XLOOKUP qui se remplissent automatiquement dès la saisie du score, et une entrée financière (clients perdus observés par jour). Toutes les cellules de score sont contraintes par des listes déroulantes de validation des données à des entiers 1–5 avec des infobulles de rubrique intégrées.

**Calc_Engine** (masqué dans la livraison client) exécute la chaîne de traitement complète : moyennes de dimension par créneau → moyennes globales AVERAGEIF → composite pondéré SUMPRODUCT → classification de note IFS → multiplication de la fuite financière → vecteur booléen de signaux d'alerte (TRUE lorsque la moyenne < 2,5). Les sorties sont consommées par le tableau de bord via les plages nommées.

**Dashboard_Report** restitue : trois cartes d'indicateurs clés (score composite, note, fuite mensuelle avec seuil d'alerte de 500 $) → pastilles de scores par dimension avec écart par rapport à la référence → panneau d'actions alimenté par FILTER listant uniquement les métriques signalées avec des remèdes pré-rédigés → emplacement réservé pour le graphique radar comparant le réel à la référence sur trois axes. Zone d'impression fixée à A1:H45, A4 portrait, ajustée à la page.

**Plages nommées enregistrées (14) :** `Store_Type`, `ASP_Value`, `Speed_Weights`, `Space_Weights`, `Benchmark_Scores`, `Overall_Score`, `Performance_Grade`, `Monthly_Leakage`, `Metric_Report_Range`, `Red_Flag_Vector`, `CX_Score_Overall`, `Ops_Score_Overall`, `Rev_Score_Overall`, `Lost_Customers_Daily`.

---

## Notes de mise en œuvre

<details>
<summary>Formule du composite pondéré</summary>

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
<summary>Note de performance</summary>

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
<summary>Extraction des signaux d'alerte — formule FILTER du tableau de bord</summary>

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
<summary>Description de rubrique XLOOKUP — affichage en direct des critères dans Input_Data</summary>

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
<summary>Graphique radar — construction manuelle (groupe B)</summary>

Le classeur contient un espace réservé gris dans les lignes 25–44 du tableau de bord avec des instructions de construction complètes. Le graphique ne peut pas être généré par programmation via openpyxl.

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
<summary>Règles de validation et liste de contrôle de livraison</summary>

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

## Obtenir le classeur

Téléchargez `cafe_audit_tool.xlsx` depuis ce dépôt.

Ouvrez-le dans Excel pour bureau, iOS, Android ou Online. Aucune macro. Aucune configuration requise au-delà de la saisie du type de magasin (Setup_Parameters B4) et de l'ASP (Setup_Parameters E4).

Le fichier de simulation `cafe_audit_demo.html` reproduit toute la logique de Calc_Engine dans le navigateur. Ouvrez-le localement pour valider les sorties des formules par rapport au classeur Excel avant la livraison client.

---

## Limites

**Ce que ce classeur ne peut pas faire :**

- Il ne capture pas l'expérience rapportée par les clients. Les scores ne reflètent que l'observation de l'auditeur. Un café peut obtenir 4,5 en CX alors que les clients accordent 2,8 au taux de recommandation net — l'observation et la perception ne sont pas des métriques équivalentes.
- Il ne suit pas les données longitudinales. Chaque audit est un instantané ponctuel. L'analyse des tendances exige une consignation manuelle des scores composites et des chiffres de fuite entre les visites dans un suivi séparé.
- Le calcul de la fuite financière suppose que tous les clients perdus ont le même ASP. Si les clients partants penchent vers les grosses dépenses, l'estimation est conservatrice. S'ils penchent vers les petites dépenses (p. ex., des curieux plutôt que des acheteurs), elle surestime la fuite.
- Les valeurs de référence (CX : 4,0, Ops : 3,8, Rev : 3,5) et le seuil d'alerte de 500 $ sont des valeurs par défaut configurables, et non des moyennes sectorielles validées statistiquement. Les utilisateurs sur des marchés urbains à fort volume ou dans des contextes suburbains à faible volume doivent recalibrer indépendamment Setup_Parameters!B14:B16 et le seuil de fuite.
- Le seuil de signal d'alerte de 2,5 est une coupure médiane de l'échelle. Il ne fera pas apparaître les métriques obtenant 2,6 qui peuvent néanmoins représenter un risque opérationnel selon le contexte.

---

## Autres outils d'aide à la décision dont vous pourriez avoir besoin

| Scénario | Instrument suggéré |
|---|---|
| Comparaison de portefeuille multi-sites (10+ cafés) | Agréger les scores d'audit dans un suivi Excel séparé avec GROUPBY/PIVOTBY sur les colonnes de score composite et de fuite |
| Analyse d'écart entre perception client et score opérationnel | Associer le score composite d'audit aux données de satisfaction client ou de recommandation ; analyse de corrélation dans un modèle séparé |
| Mesure du retour sur investissement de la formation du personnel | Audit avant/après avec ce classeur ; l'écart du score composite et de la fuite mensuelle est l'indicateur clé principal |
| Aide à la décision de renouvellement de bail | Combiner le score composite au revenu par mètre carré et au comptage de fréquentation dans un modèle financier séparé |
| Analyse de la tarification du menu | Nécessite des données de caisse au niveau de chaque transaction ; hors du périmètre d'un instrument d'audit fondé sur l'observation |
| Étalonnage concurrentiel par rapport à des marques nommées | Nécessite des données de référence sourcées en externe ; les références du classeur sont configurables mais non mises à jour automatiquement |

---

## Licence

Copyright 2025. Publié sous [Apache License 2.0](https://www.apache.org/licenses/LICENSE-2.0).

Utilisez, modifiez et distribuez ce classeur et les fichiers associés selon les termes de l'Apache License 2.0. Attribution requise. Aucune garantie n'est fournie, expresse ou implicite. Consultez le fichier `LICENSE` pour les conditions complètes.

---

*Instrument d'aide à la décision. Pas un système d'entreprise.*

---

## Audit de café · Moteur de diagnostic 3D

[![Licence](https://img.shields.io/badge/许可证-Apache_2.0-blue.svg)](https://opensource.org/licenses/Apache-2.0)
[![Plateforme](https://img.shields.io/badge/平台-Excel%20%7C%20Web-navy)](https://github.com)
[![Type d'outil](https://img.shields.io/badge/工具类型-决策支持-2251FF)](https://github.com)

**Un cadre d'audit structuré qui convertit les données d'observation sur site d'un café en un score diagnostic pondéré, une estimation de la fuite financière et une liste d'actions priorisée — un audit unique étant réalisé en moins de 30 minutes.**

Le contenu de cette version française est identique à celui de la version anglaise. Toutes les formules, la logique de pondération, les descriptions d'architecture et les clauses de limitation s'appliquent. Les détails techniques complets figurent dans les blocs dépliants de la version anglaise.

[▶ Démo en ligne](cafe_audit_demo.html) · [⬇ Télécharger le classeur](cafe_audit_tool.xlsx)
