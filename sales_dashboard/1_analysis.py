"""
Sales Performance Dashboard — Python Analysis
"""

import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import seaborn as sns
import os

# ── Config ───────────────────────────────────────────────────────────────────
INPUT  = r"C:\Users\HP\Downloads\sales_dashboard_project\sales_dashboard\ecommerce_sales_dataset.csv"
OUT    = "outputs"
os.makedirs(OUT, exist_ok=True)

PALETTE = {
    "blue":   "#378ADD",
    "green":  "#1D9E75",
    "orange": "#D85A30",
    "purple": "#7F77DD",
    "amber":  "#BA7517",
}
COLORS = list(PALETTE.values())

plt.rcParams.update({
    "font.family": "DejaVu Sans",
    "axes.spines.top": False,
    "axes.spines.right": False,
    "axes.grid": True,
    "grid.alpha": 0.25,
    "figure.dpi": 140,
})

# ── Load & Clean ─────────────────────────────────────────────────────────────
df = pd.read_csv(INPUT, parse_dates=["Order_Date"])
df = df[df["Order_Status"] != "Cancelled"].copy()
df["YearMonth"] = df["Order_Date"].dt.to_period("M")
df["YQ"] = df["Year"].astype(str) + "-" + df["Quarter"]

print("=" * 50)
print("KEY PERFORMANCE INDICATORS")
print("=" * 50)
total_rev    = df["Revenue"].sum()
total_profit = df["Profit"].sum()
total_orders = len(df)
avg_margin   = df["Profit_Margin_%"].mean()
avg_order    = df["Revenue"].mean()

print(f"  Total Revenue    : ${total_rev:>12,.2f}")
print(f"  Total Profit     : ${total_profit:>12,.2f}")
print(f"  Total Orders     : {total_orders:>13,}")
print(f"  Avg Profit Margin: {avg_margin:>12.2f}%")
print(f"  Avg Order Value  : ${avg_order:>12.2f}")
print()

# Export cleaned KPI summary
kpi_df = pd.DataFrame([{
    "Metric": k, "Value": v
} for k, v in {
    "Total Revenue ($)":    round(total_rev, 2),
    "Total Profit ($)":     round(total_profit, 2),
    "Total Orders":         total_orders,
    "Avg Profit Margin (%)":round(avg_margin, 2),
    "Avg Order Value ($)":  round(avg_order, 2),
}.items()])
kpi_df.to_csv(f"{OUT}/kpis.csv", index=False)
print("Saved: outputs/kpis.csv")

# ── 1. Quarterly Revenue & Profit Trend ──────────────────────────────────────
quarterly = (
    df.groupby(["Year", "Quarter"])[["Revenue", "Profit"]]
    .sum()
    .reset_index()
    .sort_values(["Year", "Quarter"])
)
quarterly["Label"] = quarterly["Year"].astype(str) + " " + quarterly["Quarter"]

fig, ax = plt.subplots(figsize=(13, 5))
ax.fill_between(quarterly["Label"], quarterly["Revenue"], alpha=0.12, color=PALETTE["blue"])
ax.fill_between(quarterly["Label"], quarterly["Profit"],  alpha=0.12, color=PALETTE["green"])
ax.plot(quarterly["Label"], quarterly["Revenue"], marker="o", markersize=4,
        color=PALETTE["blue"],  linewidth=2, label="Revenue")
ax.plot(quarterly["Label"], quarterly["Profit"],  marker="s", markersize=4,
        color=PALETTE["green"], linewidth=2, label="Profit", linestyle="--")
ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"${x/1000:.0f}k"))
plt.xticks(rotation=45, ha="right", fontsize=8)
ax.set_title("Quarterly Revenue & Profit Trend (2021–2024)", fontsize=13, fontweight="bold", pad=14)
ax.legend(frameon=False)
plt.tight_layout()
plt.savefig(f"{OUT}/chart1_quarterly_trend.png")
plt.close()
print("Saved: outputs/chart1_quarterly_trend.png")

# ── 2. Revenue by Region (donut) ─────────────────────────────────────────────
region = df.groupby("Region")["Revenue"].sum().sort_values(ascending=False)

fig, ax = plt.subplots(figsize=(7, 6))
wedges, texts, autotexts = ax.pie(
    region.values, labels=None, autopct="%1.1f%%",
    colors=COLORS[:len(region)], startangle=140,
    wedgeprops={"width": 0.55, "edgecolor": "white", "linewidth": 2},
    pctdistance=0.78
)
for at in autotexts:
    at.set(fontsize=9, fontweight="bold", color="white")
ax.legend(region.index, loc="lower center", ncol=2, frameon=False,
          bbox_to_anchor=(0.5, -0.08), fontsize=10)
ax.set_title("Revenue by Region", fontsize=13, fontweight="bold", pad=14)
plt.tight_layout()
plt.savefig(f"{OUT}/chart2_region_donut.png")
plt.close()
print("Saved: outputs/chart2_region_donut.png")

# ── 3. Revenue & Profit by Category (horizontal bar) ─────────────────────────
cat = (
    df.groupby("Category")[["Revenue", "Profit"]]
    .sum()
    .sort_values("Revenue")
)

fig, ax = plt.subplots(figsize=(9, 5))
y = range(len(cat))
bars_r = ax.barh(y, cat["Revenue"], color=PALETTE["blue"],  alpha=0.85, label="Revenue", height=0.45)
bars_p = ax.barh([i + 0.45 for i in y], cat["Profit"], color=PALETTE["green"], alpha=0.85, label="Profit",  height=0.45)
ax.set_yticks([i + 0.22 for i in y])
ax.set_yticklabels(cat.index)
ax.xaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"${x/1e6:.1f}M"))
for bar in bars_r:
    ax.text(bar.get_width() + 5000, bar.get_y() + bar.get_height()/2,
            f"${bar.get_width()/1000:.0f}k", va="center", fontsize=8)
ax.set_title("Revenue & Profit by Category", fontsize=13, fontweight="bold", pad=14)
ax.legend(frameon=False)
plt.tight_layout()
plt.savefig(f"{OUT}/chart3_category_bars.png")
plt.close()
print("Saved: outputs/chart3_category_bars.png")

# ── 4. Monthly Revenue Heatmap ────────────────────────────────────────────────
pivot = df.pivot_table(index="Year", columns="Month", values="Revenue", aggfunc="sum")
pivot.columns = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]

fig, ax = plt.subplots(figsize=(13, 4))
sns.heatmap(pivot, annot=True, fmt=".0f", cmap="Blues", linewidths=0.4,
            ax=ax, cbar_kws={"format": "${x:,.0f}"}, annot_kws={"size": 8})
ax.set_title("Monthly Revenue Heatmap by Year ($)", fontsize=13, fontweight="bold", pad=14)
ax.set_ylabel("")
plt.tight_layout()
plt.savefig(f"{OUT}/chart4_monthly_heatmap.png")
plt.close()
print("Saved: outputs/chart4_monthly_heatmap.png")

# ── 5. Top 10 Products by Revenue ────────────────────────────────────────────
top_products = (
    df.groupby("Product_Name")["Revenue"]
    .sum()
    .sort_values(ascending=False)
    .head(10)
)

fig, ax = plt.subplots(figsize=(10, 6))
bars = ax.barh(top_products.index[::-1], top_products.values[::-1],
               color=PALETTE["purple"], alpha=0.85)
ax.xaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"${x/1000:.0f}k"))
for bar in bars:
    ax.text(bar.get_width() + 1000, bar.get_y() + bar.get_height()/2,
            f"${bar.get_width()/1000:.0f}k", va="center", fontsize=8)
ax.set_title("Top 10 Products by Revenue", fontsize=13, fontweight="bold", pad=14)
plt.tight_layout()
plt.savefig(f"{OUT}/chart5_top_products.png")
plt.close()
print("Saved: outputs/chart5_top_products.png")

# ── 6. Customer Segment × Profit Margin scatter ───────────────────────────────
seg = (
    df.groupby("Customer_Segment")
    .agg(Revenue=("Revenue","sum"), Profit=("Profit","sum"),
         Orders=("Order_ID","count"), Margin=("Profit_Margin_%","mean"))
    .reset_index()
)

fig, ax = plt.subplots(figsize=(8, 5))
for i, row in seg.iterrows():
    ax.scatter(row["Revenue"], row["Margin"], s=row["Orders"]/3,
               color=COLORS[i], alpha=0.8, label=row["Customer_Segment"])
    ax.annotate(row["Customer_Segment"], (row["Revenue"], row["Margin"]),
                textcoords="offset points", xytext=(8, 4), fontsize=9)
ax.xaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"${x/1e6:.1f}M"))
ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"{x:.1f}%"))
ax.set_xlabel("Total Revenue")
ax.set_ylabel("Avg Profit Margin (%)")
ax.set_title("Customer Segment: Revenue vs Profit Margin\n(bubble size = order count)",
             fontsize=12, fontweight="bold", pad=14)
plt.tight_layout()
plt.savefig(f"{OUT}/chart6_segment_scatter.png")
plt.close()
print("Saved: outputs/chart6_segment_scatter.png")

print()
print("All done! Check the outputs/ folder.")
