# 📐 Power BI DAX Measures & Calculated Columns

> **Table name:** `Cleaned Dataset`
> All measures should be created by right-clicking `Cleaned Dataset` in the Fields pane → **New measure**
> All calculated columns: right-click → **New column**

---

## ✅ KPI Measures

```dax
-- Total titles
Total Titles = COUNTROWS('Cleaned Dataset')

-- Total movies
Total Movies =
CALCULATE(
    COUNTROWS('Cleaned Dataset'),
    'Cleaned Dataset'[Type] = "Movie"
)

-- Total TV shows
TV Shows =
CALCULATE(
    COUNTROWS('Cleaned Dataset'),
    'Cleaned Dataset'[Type] = "TV Show"
)

-- Total revenue
Total Revenue = SUM('Cleaned Dataset'[Estimated Revenue Usd])

-- Average revenue per title
Avg Revenue = AVERAGE('Cleaned Dataset'[Estimated Revenue Usd])

-- Average completion rate
Avg Completion Rate = AVERAGE('Cleaned Dataset'[Completion Rate Pct])

-- Average engagement score
Avg Engagement = AVERAGE('Cleaned Dataset'[Engagement Score])

-- Average watch time
Avg Watch Time = AVERAGE('Cleaned Dataset'[Avg Watch Time Pct])

-- Average movie runtime (movies only)
Avg Movie Runtime =
CALCULATE(
    AVERAGE('Cleaned Dataset'[Duration Minutes]),
    'Cleaned Dataset'[Type] = "Movie"
)

-- Distinct countries
Countries =
CALCULATE(
    DISTINCTCOUNT('Cleaned Dataset'[Country]),
    'Cleaned Dataset'[Country] <> "Not Specified"
)

-- Adult content percentage
Pct Adult Content =
DIVIDE(
    CALCULATE(
        COUNTROWS('Cleaned Dataset'),
        'Cleaned Dataset'[Rating Group] = "Adult"
    ),
    [Total Titles],
    0
) * 100

-- Recent titles percentage (added within 2 years of release)
Pct Recent Titles =
DIVIDE(
    CALCULATE(
        COUNTROWS('Cleaned Dataset'),
        'Cleaned Dataset'[Is_Recent_Release] = "Yes"
    ),
    [Total Titles],
    0
) * 100

-- Revenue per view
Revenue Per View =
DIVIDE(
    SUM('Cleaned Dataset'[Estimated Revenue Usd]),
    SUM('Cleaned Dataset'[Views]),
    0
)

-- Year-on-year change in titles added
YoY Change =
VAR CurrentYear = MAX('Cleaned Dataset'[Year Added])
VAR ThisYear =
    CALCULATE(
        [Total Titles],
        'Cleaned Dataset'[Year Added] = CurrentYear
    )
VAR LastYear =
    CALCULATE(
        [Total Titles],
        'Cleaned Dataset'[Year Added] = CurrentYear - 1
    )
RETURN ThisYear - LastYear
```

---

## 📐 Calculated Columns

> These are already pre-built in `Netflix_PowerBI_Fixed.xlsx`.
> Only use these if you are working from the original raw file.

```dax
-- Year Added (extract from text date YYYY-MM-DD)
Year Added =
IFERROR(
    VALUE(LEFT('Cleaned Dataset'[Date Added], 4)),
    BLANK()
)

-- Month number
Month Num =
IFERROR(
    VALUE(MID('Cleaned Dataset'[Date Added], 6, 2)),
    BLANK()
)

-- Month name
Month Name =
FORMAT(
    DATE(2024, VALUE(MID('Cleaned Dataset'[Date Added], 6, 2)), 1),
    "MMMM"
)

-- Quarter added
Quarter Added =
"Q" & ROUNDUP(
    VALUE(MID('Cleaned Dataset'[Date Added], 6, 2)) / 3,
    0
)

-- Release era (5-year buckets)
Release Era =
SWITCH(TRUE(),
    'Cleaned Dataset'[Release Year] >= 2000
        && 'Cleaned Dataset'[Release Year] <= 2004, "2000–2004",
    'Cleaned Dataset'[Release Year] >= 2005
        && 'Cleaned Dataset'[Release Year] <= 2009, "2005–2009",
    'Cleaned Dataset'[Release Year] >= 2010
        && 'Cleaned Dataset'[Release Year] <= 2014, "2010–2014",
    'Cleaned Dataset'[Release Year] >= 2015
        && 'Cleaned Dataset'[Release Year] <= 2019, "2015–2019",
    'Cleaned Dataset'[Release Year] >= 2020
        && 'Cleaned Dataset'[Release Year] <= 2024, "2020–2024",
    "2025+"
)

-- Duration bucket (movies only)
Duration Bucket =
SWITCH(TRUE(),
    'Cleaned Dataset'[Type] = "TV Show",             "TV Show",
    'Cleaned Dataset'[Duration Minutes] < 90,        "Under 90 min",
    'Cleaned Dataset'[Duration Minutes] < 110,       "90–110 min",
    'Cleaned Dataset'[Duration Minutes] < 130,       "110–130 min",
    'Cleaned Dataset'[Duration Minutes] < 150,       "130–150 min",
    'Cleaned Dataset'[Duration Minutes] < 170,       "150–170 min",
    "170+ min"
)

-- Rating group
Rating Group =
SWITCH(TRUE(),
    'Cleaned Dataset'[Rating] IN {"TV-MA","R","NC-17"},          "Adult",
    'Cleaned Dataset'[Rating] IN {"TV-14","PG-13"},              "Teen",
    'Cleaned Dataset'[Rating] IN {"PG","G","TV-PG","TV-G",
        "TV-Y","TV-Y7"},                                         "Family",
    "Unrated"
)
```

---

## 🔧 Common Errors & Fixes

| Error | Cause | Fix |
|---|---|---|
| `Cannot convert '18-S' to Number` | Power BI misread Date Added as `mm-dd-yy` | Use `Netflix_PowerBI_Fixed.xlsx` — Date Added is pre-forced to text |
| `Column not found` | Wrong table name (`Netflix_Data`) | Use `'Cleaned Dataset'` |
| `Year Added shows decimals on axis` | Column type is Decimal | Column tools → Data type → Whole Number |
| `Duration Bucket shows blank bar` | TV Shows have no Duration Minutes | Add visual-level filter: `Type = "Movie"` |

---

*Netflix Content Library Analysis — May 2026*
