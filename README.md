# 🎬 Netflix Content Library — Data Analysis Project

![Python](https://img.shields.io/badge/Python-3.10-blue?logo=python)
![PowerBI](https://img.shields.io/badge/Power%20BI-Dashboard-yellow?logo=powerbi)
![Excel](https://img.shields.io/badge/Excel-Export-green?logo=microsoftexcel)
![Status](https://img.shields.io/badge/Status-Completed-brightgreen)

---

## 📌 Project Overview

An end-to-end data analytics project analysing a **5,000-title Netflix content library** spanning **2015 to 2026** across **10 countries**, **9 languages**, and **12 genres**.

The goal was to uncover actionable insights around:
- Content performance and revenue generation
- Audience engagement and completion behaviour
- Geographic market opportunity
- Content strategy and commissioning decisions

---

## 🗂 Repository Structure

```
netflix-analysis/
│
├── README.md                      ← Project overview, setup & summary
├── INSIGHTS.md                    ← Key findings & stakeholder summary
├── DAX_MEASURES.md                ← All Power BI DAX measures & calculated columns
├── data_cleaning.py               ← Python data cleaning & transformation script
├── requirements.txt               ← Python dependencies
└── data/
    └── Netflix_PowerBI_Fixed.xlsx ← Cleaned, Power BI-ready dataset (25 columns)
```

---

## 📊 Dataset Summary

| Attribute | Detail |
|---|---|
| Total Titles | 5,000 |
| Movies | 2,516 (50.3%) |
| TV Shows | 2,484 (49.7%) |
| Date Range | 2015 – 2026 |
| Countries | 10 |
| Languages | 9 |
| Genres | 12 |
| Total Revenue | $2,086,040,784 |
| Avg Revenue Per Title | $417,208 |
| Avg Completion Rate | 63.8% |
| Avg Engagement Score | 41.56 |
| Avg Views Per Title | 208,383 |

---

## 🛠 Tools & Technologies

| Tool | Purpose |
|---|---|
| Python (pandas, openpyxl) | Data cleaning, transformation, calculated columns, Excel export |
| Microsoft Excel | Structured data storage and pre-aggregated summary sheets |
| Microsoft Power BI | Interactive dashboard, KPI cards, DAX measures and slicers |
| Notion | Project documentation and stakeholder reporting |
| GitHub | Version control and project documentation |

---

## 🧹 Data Cleaning Summary

| Issue Found | Resolution |
|---|---|
| `Date Added` corrupt in Power BI (`18-S` error) | Forced to plain text `@` format; all date columns pre-calculated in Python |
| Missing Director values | Filled with `"Unknown"` |
| Missing Country values | Filled with `"Not Specified"` |
| Missing Rating values | Filled with `"NR"` (Not Rated) |
| Missing Release Year values | Imputed with column median |
| Duplicate records (Title + Type) | Removed — kept first occurrence |
| Wrong table name in DAX (`Netflix_Data`) | Corrected to `Cleaned Dataset` |

---

## 📐 Calculated Columns Added (Pre-built in Python)

| Column | Description |
|---|---|
| `Year Added` | Year extracted from Date Added (2015–2026) |
| `Month Num` | Month number 1–12 for chart sorting |
| `Month Name` | Full month name (January … December) |
| `Quarter Added` | Q1, Q2, Q3, Q4 |
| `Release Era` | 5-year era bucket (2000–2004, 2005–2009 …) |
| `Duration Bucket` | Movie runtime category (Under 90 min, 90–110 min …) |
| `Rating Group` | Adult / Teen / Family / Unrated |
| `Churn Label` | Movie or TV Show label |

---

## 📈 Power BI Dashboard

### KPI Cards
Total Titles · Total Movies · Total TV Shows · Avg Movie Runtime · Countries · Adult Content % · Total Revenue · Avg Completion Rate

### Charts (12 Visuals)

| Chart | Type |
|---|---|
| Titles by genre | Horizontal bar |
| Movies vs TV Shows | Donut |
| Revenue by genre | Horizontal bar |
| Revenue by country | Column |
| Completion rate by genre | Horizontal bar |
| Views vs revenue | Scatter |
| Engagement vs watch time by genre | Clustered bar |
| Content by release era | Column |
| Movie runtime distribution | Column |
| Titles added per year | Area / Line |
| Language popularity | Horizontal bar |
| Revenue by country | Filled map |

### Slicers
Type · Genre · Rating Group · Language · Country · Content Age Bucket · Year Added

---

## 💡 Key Findings

- **France ($227.5M) outearns the USA ($197.4M)** in total revenue
- **Comedy earns the most per title** at $457K — 18% above Animation ($386K)
- **Action has the highest loyalty** — 65.5% completion rate, 42.53 engagement score
- **Views and revenue strongly correlated** (r = 0.83)
- **Portuguese content ranks #1** in popularity index (33.3)
- **Average movie runtime is 130 minutes** — above the 100-min industry average
- **2022 was the peak year** for content additions
- **January and October** are the strongest engagement months

---

## ✅ Recommendations

1. Shift investment toward **France and Japan** — highest revenue per title
2. Commission more **Comedy and Fantasy** — consistently top earners
3. Audit **Crime content** — lowest engagement, completion and revenue
4. Use **Action content** for subscriber retention campaigns
5. Expand **Portuguese and Korean originals** — #1 and #2 in popularity
6. Build a **short-film strategy** — under 90 min gap on mobile
7. Schedule major releases in **January and October** — peak months

---

## 🚀 How to Run

```bash
# 1. Clone the repo
git clone https://github.com/YOUR_USERNAME/netflix-analysis.git
cd netflix-analysis

# 2. Install dependencies
pip install -r requirements.txt

# 3. Add raw file to data/ folder as Netflix_Data.xlsx

# 4. Run the cleaning script
python data_cleaning.py
# Output → data/Netflix_PowerBI_Fixed.xlsx
```

---

*Project completed — May 2026*
