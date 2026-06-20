# BMW GROUP SALES — COMPLETE PYTHON ANALYTICS PIPELINE
# Run: pip install pandas numpy matplotlib seaborn plotly openpyxl

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import warnings

warnings.filterwarnings("ignore")

# =========================
# 1. LOAD & CLEAN DATA
# =========================

FILE_NAME = "BMW_Sales _Analytics (1).xlsx"

df = pd.read_excel(FILE_NAME, sheet_name="Original_Data", header=1)

cols = [
    'Year','Month','Region','Model','Units_Sold','Avg_Price_EUR',
    'Revenue_EUR','BEV_Share','Premium_Share','GDP_Growth',
    'Fuel_Price_Index'
]

df = df[cols].copy()

df['BEV_Pct'] = df['BEV_Share'] * 100
df['Month_Name'] = pd.to_datetime(df['Month'], format='%m').dt.strftime('%b')

df['Segment'] = df['Model'].map({
    '3 Series':'Entry Premium',
    '5 Series':'Mid Premium',
    'MINI':'Entry Premium',
    'X3':'Mid Premium',
    'X5':'Mid Premium',
    'X7':'Luxury',
    'iX':'Electric'
})

print("=" * 60)
print("BMW GROUP SALES ANALYTICS")
print("=" * 60)

# =========================
# 2. KPI SUMMARY
# =========================

total_revenue = df['Revenue_EUR'].sum() / 1e9
total_units = df['Units_Sold'].sum() / 1e6
avg_bev = df['BEV_Pct'].mean()

print(f"Total Revenue : EUR {total_revenue:.2f} B")
print(f"Total Units   : {total_units:.2f} M")
print(f"Average BEV%  : {avg_bev:.2f}%")

# =========================
# 3. AGGREGATIONS
# =========================

rev_year = df.groupby('Year')['Revenue_EUR'].sum() / 1e9
unit_year = df.groupby('Year')['Units_Sold'].sum()

rev_region = df.groupby('Region')['Revenue_EUR'].sum() / 1e9
rev_model = df.groupby('Model')['Revenue_EUR'].sum().sort_values(ascending=False) / 1e9

bev_year = df.groupby('Year')['BEV_Pct'].mean()

# =========================
# 4. EXECUTIVE INSIGHTS
# =========================

print("\n===== EXECUTIVE INSIGHTS =====")

print(f"Top Revenue Model  : {rev_model.idxmax()}")
print(f"Top Revenue Region : {rev_region.idxmax()}")
print(f"Average BEV Share  : {bev_year.mean():.1f}%")

# =========================
# 5. DASHBOARD (MATPLOTLIB)
# =========================

plt.style.use("seaborn-v0_8-whitegrid")

fig, axes = plt.subplots(2, 3, figsize=(20, 11))

fig.suptitle(
    "BMW Group Sales Intelligence Dashboard",
    fontsize=16,
    fontweight="bold"
)

axes[0,0].bar(rev_year.index, rev_year.values)
axes[0,0].set_title("Revenue by Year")

axes[0,1].pie(
    rev_region.values,
    labels=rev_region.index,
    autopct="%1.1f%%"
)
axes[0,1].set_title("Revenue Share by Region")

axes[0,2].barh(rev_model.index, rev_model.values)
axes[0,2].set_title("Revenue by Model")

axes[1,0].plot(
    bev_year.index,
    bev_year.values,
    marker="o",
    linewidth=2
)
axes[1,0].set_title("BEV Share Trend")

pivot_region = df.pivot_table(
    values="Revenue_EUR",
    index="Year",
    columns="Region",
    aggfunc="sum"
) / 1e9

pivot_region.plot(kind="bar", ax=axes[1,1])
axes[1,1].set_title("Revenue by Region per Year")

unit_model = df.groupby("Model")["Units_Sold"].sum().sort_values()

axes[1,2].barh(
    unit_model.index,
    unit_model.values / 1e6
)
axes[1,2].set_title("Units Sold by Model")

plt.tight_layout()
plt.savefig("bmw_dashboard.png", dpi=300)

# =========================
# 6. CORRELATION MATRIX
# =========================

corr_cols = [
    'Units_Sold',
    'Avg_Price_EUR',
    'Revenue_EUR',
    'BEV_Pct',
    'GDP_Growth',
    'Fuel_Price_Index'
]

corr = df[corr_cols].corr()

plt.figure(figsize=(10, 8))

sns.heatmap(
    corr,
    annot=True,
    cmap="coolwarm",
    fmt=".2f"
)

plt.title("BMW Correlation Matrix")
plt.tight_layout()
plt.savefig("correlation_matrix.png", dpi=300)

# =========================
# 7. INTERACTIVE DASHBOARD
# =========================

fig = make_subplots(
    rows=2,
    cols=2,
    subplot_titles=(
        "Revenue by Year",
        "Revenue by Region",
        "Revenue by Model",
        "BEV Trend"
    ),
    specs=[
        [{"type":"bar"},{"type":"pie"}],
        [{"type":"bar"},{"type":"scatter"}]
    ]
)

fig.add_trace(
    go.Bar(x=list(rev_year.index), y=list(rev_year.values)),
    row=1, col=1
)

fig.add_trace(
    go.Pie(labels=list(rev_region.index),
           values=list(rev_region.values)),
    row=1, col=2
)

fig.add_trace(
    go.Bar(
        x=list(rev_model.values),
        y=list(rev_model.index),
        orientation="h"
    ),
    row=2, col=1
)

fig.add_trace(
    go.Scatter(
        x=list(bev_year.index),
        y=list(bev_year.values),
        mode="lines+markers"
    ),
    row=2, col=2
)

fig.update_layout(
    title="BMW Interactive Dashboard",
    height=900,
    showlegend=False
)

fig.write_html("bmw_interactive_dashboard.html")

# =========================
# 8. YoY ANALYSIS
# =========================

yoy = rev_year.pct_change() * 100

print("\n===== YoY Revenue Growth =====")

for year, growth in yoy.dropna().items():
    print(f"{year}: {growth:+.2f}%")

# =========================
# 9. SEGMENT ANALYSIS
# =========================

seg_rev = df.groupby("Segment")["Revenue_EUR"].sum() / 1e9

print("\n===== Revenue by Segment =====")
print(seg_rev.sort_values(ascending=False))

# =========================
# 10. EXPORT REPORT
# =========================

with pd.ExcelWriter("BMW_Analytics_Report.xlsx") as writer:
    rev_year.to_frame("Revenue_Billion").to_excel(
        writer,
        sheet_name="Revenue_Year"
    )

    rev_region.to_frame("Revenue_Billion").to_excel(
        writer,
        sheet_name="Revenue_Region"
    )

    rev_model.to_frame("Revenue_Billion").to_excel(
        writer,
        sheet_name="Revenue_Model"
    )

    seg_rev.to_frame("Revenue_Billion").to_excel(
        writer,
        sheet_name="Segments"
    )

print("\nAnalysis Complete.")
