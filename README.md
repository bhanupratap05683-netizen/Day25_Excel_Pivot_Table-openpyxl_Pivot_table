# Day25_Excel_Pivot_Table-openpyxl_Pivot_table 

**84-Day Python & Excel Roadmap** · Phase 2 · Day 25 of 84

---

## What I Built

Loaded 60 rows of financial sales data, built pivot-style summaries using Python dictionaries and openpyxl, and exported formatted results to Excel.

---

## Files

| File | Description |
|------|-------------|
| `Day25_PivotTables_Practice.xlsx` | Input — 60 rows of sales data (5 regions, 5 products, 2 quarters) |
| `day25_pivot_analysis.py` | Script — reads Excel, groups data, writes pivot summaries |
| `Day25_Pivot_Output.xlsx` | Output — 6 pivot summaries across 2 sheets |

---

## Concepts Practiced

**Excel**
- Pivot Tables — Rows, Columns, Values, Filters zones
- Slicers — visual filter buttons for dashboards
- Calculated Fields — custom formula columns inside a pivot
- Value Field Settings — switch SUM to AVERAGE per field
- 2D Pivot — Region × Quarter cross-tab view

**Python (openpyxl)**
- Read all rows into a list of dictionaries via `iter_rows`
- Group and sum using `dict.get(key, 0)` pattern
- Compute averages with a separate count dictionary
- Build 2D cross-tab using nested dictionaries
- Write all summaries to a formatted output Excel file

---

## Key Insight

Pivot Tables are just grouped dictionary aggregations under the hood:

```python
for row in rows:
    key = row["Region"]
    region_revenue[key] = region_revenue.get(key, 0) + row["Revenue"]
```

This is exactly what dragging Region → Rows and Revenue → Values (SUM) does in Excel.

---

## How to Run

```bash
pip install openpyxl
python day25_pivot_analysis.py
```

---

**Bhanu Pratap Singh**  · [GitHub](https://github.com/YOUR_USERNAME)
