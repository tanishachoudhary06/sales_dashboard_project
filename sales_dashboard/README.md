# Sales Performance Dashboard — Project

E-Commerce Sales Analysis · 10,000 Orders · 2021–2024

## Files

| File | Purpose |
|---|---|
| `1_analysis.py` | Python charts (matplotlib/seaborn) — 6 PNG charts + KPI CSV |
| `2_excel_dashboard.py` | Excel workbook with 5 sheets and embedded charts |
| `3_web_dashboard.html` | Standalone web dashboard (Chart.js, no server needed) |
| `4_tableau_prep.py` | Exports 7 clean CSVs for Tableau data sources |
| `ecommerce_sales_dataset.csv` | Source data |

## Quick Start

### 1 · Python Charts
```bash
pip install pandas matplotlib seaborn
python 1_analysis.py
# → outputs/chart1_quarterly_trend.png … chart6_segment_scatter.png
```

### 2 · Excel Dashboard
```bash
pip install pandas openpyxl
python 2_excel_dashboard.py
# → outputs/Sales_Dashboard.xlsx  (5 sheets with charts)
```

### 3 · Web Dashboard
Open `3_web_dashboard.html` in a browser — **must be served locally**:
```bash
# Python built-in server (from project folder)
python -m http.server 8080
# Then open: http://localhost:8080/3_web_dashboard.html
```
Features: 5 tabs (Overview, Regional, Products, Time Series, Data Table), filters, live search.

### 4 · Tableau Prep
```bash
pip install pandas
python 4_tableau_prep.py
# → outputs/tableau/01_fact_orders.csv … 07_customer_segments.csv
```
Connect each CSV as a separate data source in Tableau Desktop.

## Dashboard KPIs
| Metric | Value |
|---|---|
| Total Revenue | $5.28M |
| Total Profit | $1.44M |
| Total Orders | 10,000 |
| Avg Profit Margin | 23.65% |

## Key Insights
- **Electronics** dominates at 64% of revenue ($3.38M)
- **All 4 regions** are balanced (~25% each)
- Revenue peaked in **Q3 2023** ($629k)
- **Clothing** has the highest profit margin (24.8%)
- Top product: **Lenovo ThinkPad X1** ($379k)
