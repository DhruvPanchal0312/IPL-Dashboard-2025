# IPL Dashboard 2025 🏏

A data modeling and dashboard project for the **Indian Premier League (IPL) 2025** season.

---

## 📊 Dataset & Schema Overview

The dataset is structured following dimensional modeling (Star Schema) principles with dimension and fact tables:

- **`dim_team`**: Team master details, codes, franchise information, and metadata.
- **`dim_player`**: Player profiles, roles (Batsman, Bowler, All-Rounder, Wicketkeeper), nationalities, and batting/bowling styles.
- **`dim_venue`**: Stadium details, cities, countries, and pitch/ground characteristics.
- **`fact_match`**: Match-level granular data including scores, winners, margins, toss decisions, and venues.
- **`fact_player`**: Player-level performance metrics per match (runs, wickets, strike rate, economy, catches, etc.).

---

## 📁 Repository Structure

```text
├── IPL 2025.xlsx       # Primary workbook with dimension & fact tables
├── .gitignore          # Git ignore rules for office and system files
└── README.md           # Project documentation
```

---

## 🚀 Use Cases & Analysis

- **Interactive Dashboards**: Power BI / Tableau / Excel reporting.
- **Player & Team Performance Analysis**: Form tracking, head-to-head records, strike rates, economy rates.
- **Match Insights**: Toss decision impact, venue-specific trends, chasing vs defending patterns.

---

## 👤 Author

- **Dhruv Panchal** ([@DhruvPanchal0312](https://github.com/DhruvPanchal0312))
