"""
Sales Performance Dashboard — Tableau Data Prep
Exports 6 clean CSV files ready to connect as Tableau data sources
Usage: python 4_tableau_prep.py
"""

import pandas as pd
import os

INPUT = r"C:\Users\HP\Downloads\sales_dashboard_project\sales_dashboard\ecommerce_sales_dataset.csv"
OUT   = "outputs/tableau"
os.makedirs(OUT, exist_ok=True)

df = pd.read_csv(INPUT, parse_dates=["Order_Date"])
df = df[df["Order_Status"] != "Cancelled"].copy()
df["YearMonth"] = df["Order_Date"].dt.to_period("M").astype(str)
df["YQ"]        = df["Year"].astype(str) + "-" + df["Quarter"]

# ── 1. Master fact table (Tableau main source) ────────────────────────────────
fact = df[[
    "Order_ID","Order_Date","Year","Month","Quarter","Season",
    "Customer_ID","Customer_Gender","Customer_Segment",
    "Region","Country","Category","Sub_Category","Product_Name",
    "Unit_Price","Quantity","Discount","Revenue","Cost","Profit",
    "Profit_Margin_%","Shipping_Cost","Shipping_Method",
    "Shipping_Days","Payment_Method","Order_Status"
]].copy()
fact.columns = [c.replace("%","Pct") for c in fact.columns]
fact.to_csv(f"{OUT}/01_fact_orders.csv", index=False)
print("Saved: 01_fact_orders.csv  — master fact table")

# ── 2. KPI summary ────────────────────────────────────────────────────────────
kpi = pd.DataFrame([
    {"Metric": "Total Revenue ($)",     "Value": round(df["Revenue"].sum(), 2)},
    {"Metric": "Total Profit ($)",      "Value": round(df["Profit"].sum(), 2)},
    {"Metric": "Total Orders",          "Value": len(df)},
    {"Metric": "Avg Profit Margin (%)", "Value": round(df["Profit_Margin_%"].mean(), 2)},
    {"Metric": "Avg Order Value ($)",   "Value": round(df["Revenue"].mean(), 2)},
    {"Metric": "Avg Discount",          "Value": round(df["Discount"].mean(), 4)},
    {"Metric": "Unique Customers",      "Value": df["Customer_ID"].nunique()},
    {"Metric": "Unique Products",       "Value": df["Product_Name"].nunique()},
])
kpi.to_csv(f"{OUT}/02_kpis.csv", index=False)
print("Saved: 02_kpis.csv         — KPI metrics")

# ── 3. Quarterly time series ──────────────────────────────────────────────────
quarterly = (
    df.groupby(["Year","Quarter","YQ"])
    .agg(
        Revenue      =("Revenue","sum"),
        Profit       =("Profit","sum"),
        Orders       =("Order_ID","count"),
        Avg_Margin   =("Profit_Margin_%","mean"),
        Avg_Discount =("Discount","mean"),
    )
    .reset_index()
    .sort_values(["Year","Quarter"])
)
quarterly["Profit_Margin"] = (quarterly["Profit"] / quarterly["Revenue"] * 100).round(2)
quarterly.to_csv(f"{OUT}/03_quarterly_trend.csv", index=False)
print("Saved: 03_quarterly_trend.csv  — time series (quarterly)")

# ── 4. Monthly time series ────────────────────────────────────────────────────
monthly = (
    df.groupby(["Year","Month","YearMonth"])
    .agg(
        Revenue =("Revenue","sum"),
        Profit  =("Profit","sum"),
        Orders  =("Order_ID","count"),
    )
    .reset_index()
    .sort_values(["Year","Month"])
)
monthly["Profit_Margin"] = (monthly["Profit"] / monthly["Revenue"] * 100).round(2)
monthly.to_csv(f"{OUT}/04_monthly_trend.csv", index=False)
print("Saved: 04_monthly_trend.csv    — time series (monthly)")

# ── 5. Region × Category matrix ───────────────────────────────────────────────
region_cat = (
    df.groupby(["Region","Category","Sub_Category"])
    .agg(
        Revenue      =("Revenue","sum"),
        Profit       =("Profit","sum"),
        Orders       =("Order_ID","count"),
        Avg_Margin   =("Profit_Margin_%","mean"),
        Avg_Discount =("Discount","mean"),
    )
    .reset_index()
)
region_cat["Profit_Margin"] = (region_cat["Profit"] / region_cat["Revenue"] * 100).round(2)
total_rev = region_cat["Revenue"].sum()
region_cat["Revenue_Share_Pct"] = (region_cat["Revenue"] / total_rev * 100).round(2)
region_cat.to_csv(f"{OUT}/05_region_category.csv", index=False)
print("Saved: 05_region_category.csv  — region × category breakdown")

# ── 6. Product performance ────────────────────────────────────────────────────
products = (
    df.groupby(["Product_Name","Category","Sub_Category"])
    .agg(
        Revenue      =("Revenue","sum"),
        Profit       =("Profit","sum"),
        Orders       =("Order_ID","count"),
        Units_Sold   =("Quantity","sum"),
        Avg_Unit_Price=("Unit_Price","mean"),
        Avg_Discount =("Discount","mean"),
        Avg_Margin   =("Profit_Margin_%","mean"),
    )
    .reset_index()
    .sort_values("Revenue", ascending=False)
)
products["Profit_Margin"] = (products["Profit"] / products["Revenue"] * 100).round(2)
products["Revenue_Rank"]  = products["Revenue"].rank(ascending=False).astype(int)
products.to_csv(f"{OUT}/06_product_performance.csv", index=False)
print("Saved: 06_product_performance.csv — product-level metrics")

# ── 7. Customer segment analysis ──────────────────────────────────────────────
customers = (
    df.groupby(["Customer_Segment","Customer_Gender","Region"])
    .agg(
        Revenue    =("Revenue","sum"),
        Profit     =("Profit","sum"),
        Orders     =("Order_ID","count"),
        Customers  =("Customer_ID","nunique"),
        Avg_Margin =("Profit_Margin_%","mean"),
    )
    .reset_index()
)
customers["Revenue_Per_Customer"] = (customers["Revenue"] / customers["Customers"]).round(2)
customers.to_csv(f"{OUT}/07_customer_segments.csv", index=False)
print("Saved: 07_customer_segments.csv — customer segment analysis")

# ── Summary ───────────────────────────────────────────────────────────────────
print()
print("=" * 58)
print("TABLEAU SETUP GUIDE")
print("=" * 58)
print()
print("Recommended Tableau Workbook Structure:")
print()
print("  Dashboard 1 — Executive Overview")
print("    Source: 02_kpis.csv (KPI tiles)")
print("    Source: 03_quarterly_trend.csv (line chart)")
print("    Source: 05_region_category.csv (map + donut)")
print()
print("  Dashboard 2 — Regional Deep Dive")
print("    Source: 05_region_category.csv")
print("    Viz: Filled map (Revenue by Country/Region)")
print("    Viz: Bar chart (Region vs Category heatmap)")
print()
print("  Dashboard 3 — Product Analysis")
print("    Source: 06_product_performance.csv")
print("    Viz: Treemap (Category → Sub-Category → Product)")
print("    Viz: Scatter (Revenue vs Margin, size=Orders)")
print()
print("  Dashboard 4 — Time Intelligence")
print("    Source: 04_monthly_trend.csv")
print("    Viz: Line chart with year-over-year comparison")
print("    Source: 03_quarterly_trend.csv")
print("    Viz: Bar chart (seasonal breakdown)")
print()
print("  Dashboard 5 — Customer Segments")
print("    Source: 07_customer_segments.csv")
print("    Viz: Bubble chart + heatmap by segment × region")
print()
print("All CSV files saved to: outputs/tableau/")
