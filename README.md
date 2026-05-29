# Acquisition-de-leads-inbound-outbound
Projet data analyst portfolio simulant une stratégie complète de génération de leads inbound pour une PME de rachat de véhicules d'occasion (France, 2025). 
# 🚗 V Market — Stratégie Inbound Lead Generation · Analyse Data Complète

> **Projet portfolio Data Analyst** · Simulation réaliste basée sur benchmarks marché  
> Secteur : Rachat de véhicules d'occasion · France · Année 2025

---

## 📌 Contexte du projet

**V Market** est une PME spécialisée dans le rachat de véhicules auprès de particuliers.

En janvier 2025, l'entreprise génère **100% de ses leads via des canaux outbound** (scraping d'annonces sur Le Bon Coin et La Centrale via une extension CRM). Ce modèle fonctionne en volume, mais présente trois risques structurels :

- **Dépendance critique** à deux plateformes tierces
- **Coût d'acquisition caché** (temps humain non mesuré = ~8 €/lead)
- **Zéro lead entrant** : aucun vendeur ne trouve l'entreprise par lui-même

**Mission data** : concevoir et mesurer une stratégie de génération de leads inbound à faible coût, en suivre la montée en puissance, et démontrer son impact sur la rentabilité du funnel.

---

## 🎯 Objectifs business

| Objectif | Métrique cible | Horizon |
|---|---|---|
| Réduire la dépendance outbound | Part inbound > 20% | 12 mois |
| Améliorer le taux de conversion | Closing inbound > 22% | 6 mois |
| Réduire le CPA global | CPA < 100 € | 12 mois |
| Diversifier les sources | ≥ 4 canaux actifs | 6 mois |

---

## 📊 Stack & outils

```
Python 3.11       pandas · numpy · matplotlib · seaborn
Excel / openpyxl  Workbook multi-onglets avec KPIs dynamiques
Power BI / Looker Dashboard interactif (connexion au dataset CSV)
CRM (BWA)         Tracking source obligatoire sur chaque lead
GitHub Actions    (optionnel) Re-génération automatique du dataset
```

---

## 📁 Structure du projet

```
v-market-analytics/
│
├── data/
│   └── leads_raw.csv              # Dataset simulé · 1 019 leads · 2025
│
├── scripts/
│   ├── generate_dataset.py        # Génération du dataset avec ramp-up inbound
│   ├── build_analysis.py          # Workbook Excel multi-onglets
│   └── generate_charts.py         # 4 visualisations Python
│
├── dashboards/
│   ├── v_market_analytics_2025.xlsx   # Workbook complet (7 onglets)
│   ├── fig1_evolution_inbound_outbound.png
│   ├── fig2_performance_par_source.png
│   ├── fig3_funnel_comparatif.png
│   └── fig4_roi_cpa_analyse.png
│
└── README.md
```

---

## 🔢 Dataset — Dictionnaire des variables

| Variable | Type | Description |
|---|---|---|
| `lead_id` | string | Identifiant unique du lead (VM00001…) |
| `date` | date | Date de réception du lead |
| `month` | string | Mois au format YYYY-MM |
| `source` | catégorie | Canal d'origine (LeBonCoin, LaCentrale, GBP, Facebook, Formulaire, Partenariat) |
| `inbound` | boolean | True si lead entrant (initié par le prospect) |
| `city` | string | Ville de localisation du prospect |
| `brand` | string | Marque du véhicule à vendre |
| `status` | catégorie | Statut CRM : NRP · En cours · RDV planifié · RDV effectué · Signé · Perdu |
| `delay_days` | int | Délai moyen de traitement (jours) |
| `cpl_eur` | float | Coût par lead estimé (€) |
| `vehicle_value` | float | Valeur estimée du véhicule (deals signés uniquement) |
| `margin_eur` | float | Marge brute réalisée (deals signés uniquement) |
| `margin_pct` | float | Marge en % de la valeur véhicule |

---

## 📈 Résultats clés (dataset 2025)

### Vue d'ensemble

```
Total leads générés     : 1 019
Deals signés            : ~130
Taux de closing global  : ~12.8%
Part inbound (Déc 2025) : ~22%  (vs ~2% en Jan)
Marge brute totale      : ~62 000 €
```

### Comparaison Inbound vs Outbound

| Métrique | Outbound | Inbound | Avantage inbound |
|---|---|---|---|
| Taux de qualification | ~30% | ~54% | +80% |
| Taux de closing | ~13.5% | ~24% | +78% |
| Délai moyen traitement | ~6.8 jours | ~4.1 jours | −40% |
| CPA estimé | ~120–140 € | ~0–20 € | ×6–∞ |
| ROI | ~6× | ∞ (canaux 0€) | — |

### Performance par canal

| Canal | Type | Taux closing | CPA | Statut |
|---|---|---|---|---|
| Partenariat | Inbound | ~30% | 0 € | ⭐ Meilleur ROI |
| GBP | Inbound | ~27% | 0 € | ⭐ Fort potentiel |
| Formulaire web | Inbound | ~25% | ~20 € | 🟢 Excellent |
| Facebook | Inbound | ~20% | 0 € | 🟢 Bon |
| LeBonCoin | Outbound | ~14% | ~115 € | 🟡 Fonctionnel |
| La Centrale | Outbound | ~13% | ~120 € | 🟡 Fonctionnel |

---

## 📉 Analyse du Funnel global

```
Leads reçus          1 019  ████████████████████ 100%
Contactés              ~830  ████████████████░░░░  81%
Qualifiés              ~470  █████████░░░░░░░░░░░  46%
RDV planifié           ~270  █████░░░░░░░░░░░░░░░  27%
RDV effectué           ~195  ████░░░░░░░░░░░░░░░░  19%
Signés                 ~130  ██░░░░░░░░░░░░░░░░░░  13%
```

---

## 🗺️ Stratégie inbound mise en place

### Phase 1 — Quick wins (Jan–Mar 2025)
- Optimisation Google Business Profile (catégories, photos, Q&A, collecte d'avis)
- Activation de la Page Facebook pour les annonces de rachat + Marketplace
- Mise en place d'un formulaire d'estimation (Tally / Google Forms)
- Tag source obligatoire dans le CRM sur chaque lead entrant

### Phase 2 — Infrastructure (Avr–Jun 2025)
- Landing page SEO local dédiée au rachat de véhicules
- Rythme de publication sociale (TikTok + Instagram Reels : 3×/semaine)
- 10 partenariats locaux (garages, centres de contrôle technique)
- Automatisation formulaire → CRM via Zapier

### Phase 3 — Optimisation (Jul–Déc 2025)
- Analyse des canaux les plus performants (data-driven)
- Renforcement des canaux Partenariat et GBP
- Tableaux de bord mensuels + rapport de performance par source

---

## 📊 Visualisations

| Graphique | Description |
|---|---|
| `fig1_evolution_inbound_outbound.png` | Évolution mensuelle + % inbound sparkline |
| `fig2_performance_par_source.png` | Bubble chart qualification × closing × volume |
| `fig3_funnel_comparatif.png` | Funnel horizontal Inbound vs Outbound |
| `fig4_roi_cpa_analyse.png` | CPA par canal + Marge vs Coût |

---

## 🚀 Reproduire le projet

```bash
# Clone
git clone https://github.com/[username]/vroom-market-analytics
cd vroom-market-analytics

# Dépendances
pip install pandas numpy matplotlib openpyxl

# Générer le dataset
python scripts/generate_dataset.py

# Générer le workbook Excel
python scripts/build_analysis.py

# Générer les visualisations
python scripts/generate_charts.py
```

---

## 💡 KPI de pilotage recommandés

Suivre mensuellement dans le CRM :

| KPI | Formule | Fréquence |
|---|---|---|
| Part leads inbound | leads_inbound / total_leads | Mensuel |
| Taux de closing par canal | closed / total_leads [par source] | Mensuel |
| CPA (Coût par acquisition) | coût_total_canal / deals_signés | Mensuel |
| Délai moyen de traitement | avg(delay_days) | Mensuel |
| Taux de qualification | leads_qualifiés / total_leads | Mensuel |
| Marge brute par canal | sum(margin_eur) [par source] | Mensuel |

---

## ⚠️ Avertissement

Ce projet est basé sur des **données entièrement simulées** à des fins de démonstration portfolio. Les volumes, chiffres et noms sont fictifs mais construits à partir de **benchmarks réels** du marché (HubSpot Lead Generation Report, Salesforce State of Sales, études secteur automobile VO France).

---

## 👤 Auteur

**Tening THIARE** — Data Analyst  

---

*Projet réalisé dans le cadre d'un portfolio GitHub démontrant des compétences en data analysis, growth analytics et business intelligence appliquées au secteur automobile.*
