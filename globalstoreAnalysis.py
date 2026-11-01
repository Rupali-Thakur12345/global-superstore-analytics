
# ==========================================================
# GLOBAL SUPERSTORE DATA ANALYTICS PROJECT
# ==========================================================

# STEP 1: Import Libraries

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os

# STEP 2: Load Dataset

file_path = r"C:\Users\user\Downloads\archive\superstore.csv"

if not os.path.exists(file_path):
    raise FileNotFoundError(f"Dataset not found: {file_path}")

df = pd.read_csv(file_path, encoding="latin1")

# Standardize column names

df.columns = (
    df.columns
    .str.strip()
    .str.lower()
    .str.replace('.', '_', regex=False)
)

# Rename unknown column

df.rename(
    columns={'è®å½æ°': 'record_count'},
    inplace=True
)

print("\nUpdated Column Names:")
print(df.columns.tolist())

print("\nFirst 5 Rows:")
print(df.head())

# Dataset information
print("\nDataset Information:")
print(df.info())

print("\nMissing Values:")
print(df.isnull().sum())


# ==========================================================
# STEP 3: DATA CLEANING & PREPARATION
# ==========================================================

# 3.1 Remove duplicate records

print("\nBefore removing duplicates:", df.shape)

df.drop_duplicates(inplace=True)

print("After removing duplicates:", df.shape)


# 3.2 Check and convert data types

print("\nOriginal Data Types:")
print(df.dtypes)

# Convert date columns into datetime format

df['order_date'] = pd.to_datetime(
    df['order_date'], errors='coerce'
)

df['ship_date'] = pd.to_datetime(
    df['ship_date'], errors='coerce'
)

# Convert numerical columns

numeric_columns = [
    'sales', 'profit', 'discount',
    'quantity', 'shipping_cost'
]

for col in numeric_columns:
    df[col] = pd.to_numeric(df[col], errors='coerce')


# 3.3 Drop unnecessary columns

columns_to_drop = ['row_id', 'ji_lu_shu']

df.drop(
    columns=columns_to_drop,
    errors='ignore',
    inplace=True
)

# 3.4 Check final structure

print("\nFinal Dataset Structure:")
print(df.info())

print("\nMissing Values After Cleaning:")
print(df.isnull().sum())


# ==========================================================
# STEP 4: EXPLORATORY DATA ANALYSIS (EDA)
# ==========================================================

# 4.1 Total Sales by Category

sales_by_category = df.groupby('category')['sales'].sum()

print("\nTotal Sales by Category:")
print(sales_by_category)


# Total Sales by Sub-category

sales_by_sub_category = (
    df.groupby('sub_category')['sales']
    .sum()
    .sort_values(ascending=False)
)

print("\nTotal Sales by Sub-category:")
print(sales_by_sub_category)


# 4.2 Most Profitable Categories and Sub-categories

profit_by_category = df.groupby('category')['profit'].sum()

print("\nTotal Profit by Category:")
print(profit_by_category)


# Profit by Sub-category

profit_by_sub_category = (
    df.groupby('sub_category')['profit']
    .sum()
    .sort_values(ascending=False)
)

print("\nTotal Profit by Sub-category:")
print(profit_by_sub_category)


# 4.3 Region-wise Sales and Profit

region_summary = (
    df.groupby('region')[['sales', 'profit']]
    .sum()
    .sort_values(by='sales', ascending=False)
)

print("\nRegion-wise Sales and Profit:")
print(region_summary)


# 4.4 Market-wise Sales and Profit

market_summary = (
    df.groupby('market')[['sales', 'profit']]
    .sum()
    .sort_values(by='sales', ascending=False)
)

print("\nMarket-wise Sales and Profit:")
print(market_summary)


# 4.5 Monthly Sales Trend

df['order_month'] = df['order_date'].dt.to_period('M')

monthly_sales = (
    df.groupby('order_month')['sales']
    .sum()
    .sort_index()
)

print("\nMonthly Sales Trend:")
print(monthly_sales.tail(12))


# 4.6 Top 10 Customers by Sales

top_customers = (
    df.groupby('customer_name')['sales']
    .sum()
    .sort_values(ascending=False)
    .head(10)
)

print("\nTop 10 Customers by Sales:")
print(top_customers)


# 4.7 Sales by Segment

segment_sales = df.groupby('segment')['sales'].sum()

print("\nSales by Customer Segment:")
print(segment_sales)


# 4.8 Sales by Ship Mode

shipping_summary = (
    df.groupby('ship_mode')[['sales', 'profit', 'shipping_cost']]
    .sum()
)

print("\nShipping Mode Analysis:")
print(shipping_summary)


# ==========================================================
# STEP 5: DATA VISUALIZATION
# ==========================================================



import matplotlib.ticker as mtick

# Professional chart theme
sns.set_theme(style="whitegrid", context="notebook")

# Consistent colors
primary_color = "#264653"
secondary_color = "#2A9D8F"
accent_color = "#E9C46A"

plt.rcParams["figure.facecolor"] = "white"
plt.rcParams["axes.facecolor"] = "white"
plt.rcParams["axes.titleweight"] = "bold"
plt.rcParams["axes.titlesize"] = 15
plt.rcParams["axes.labelsize"] = 11


# ----------------------------------------------------------
# 1. SALES BY CATEGORY
# ----------------------------------------------------------

sales_by_category = (
    df.groupby("category")["sales"]
    .sum()
    .sort_values(ascending=True)
)

fig, ax = plt.subplots(figsize=(10, 5))

bars = ax.barh(
    sales_by_category.index,
    sales_by_category.values,
    color=[secondary_color, primary_color, accent_color]
)

ax.set_title("Total Sales by Category", pad=18)
ax.set_xlabel("Total Sales")
ax.set_ylabel("Product Category")

ax.xaxis.set_major_formatter(
    mtick.FuncFormatter(lambda x, _: f"{x/1e6:.1f}M")
)

ax.bar_label(
    bars,
    labels=[f"{x/1e6:.2f}M" for x in sales_by_category.values],
    padding=5,
    fontsize=10
)

sns.despine()
plt.tight_layout()
plt.show()


# ----------------------------------------------------------
# 2. MONTHLY SALES TREND
# ----------------------------------------------------------

df["order_month"] = df["order_date"].dt.to_period("M")

monthly_sales = (
    df.groupby("order_month")["sales"]
    .sum()
    .sort_index()
)

fig, ax = plt.subplots(figsize=(13, 5))

x = monthly_sales.index.to_timestamp()

ax.plot(
    x,
    monthly_sales.values,
    marker="o",
    linewidth=2.5,
    markersize=5,
    color=secondary_color
)

ax.fill_between(
    x,
    monthly_sales.values,
    alpha=0.12,
    color=secondary_color
)

ax.set_title("Monthly Sales Performance", pad=18)
ax.set_xlabel("Month")
ax.set_ylabel("Sales")

ax.yaxis.set_major_formatter(
    mtick.FuncFormatter(lambda y, _: f"{y/1e6:.1f}M")
)

ax.grid(axis="y", alpha=0.3)
ax.grid(axis="x", visible=False)

sns.despine()
plt.tight_layout()
plt.show()


# ----------------------------------------------------------
# 3. TOP 10 SUB-CATEGORIES BY SALES
# ----------------------------------------------------------

top_subcats = (
    df.groupby("sub_category")["sales"]
    .sum()
    .nlargest(10)
    .sort_values()
)

fig, ax = plt.subplots(figsize=(10, 6))

bars = ax.barh(
    top_subcats.index,
    top_subcats.values,
    color=secondary_color
)

ax.set_title("Top 10 Sub-Categories by Sales", pad=18)
ax.set_xlabel("Total Sales")
ax.set_ylabel("Sub-Category")

ax.xaxis.set_major_formatter(
    mtick.FuncFormatter(lambda x, _: f"{x/1e6:.1f}M")
)

ax.bar_label(
    bars,
    labels=[f"{x/1e6:.2f}M" for x in top_subcats.values],
    padding=4
)

sns.despine()
plt.tight_layout()
plt.show()


# ----------------------------------------------------------
# 4. REGION-WISE SALES AND PROFIT
# ----------------------------------------------------------

region_summary = (
    df.groupby("region")[["sales", "profit"]]
    .sum()
    .sort_values("sales", ascending=False)
)

fig, ax = plt.subplots(figsize=(12, 6))

region_summary.plot(
    kind="bar",
    ax=ax,
    color=[primary_color, secondary_color],
    width=0.75
)

ax.set_title("Regional Sales vs Profit", pad=18)
ax.set_xlabel("Region")
ax.set_ylabel("Amount")

ax.yaxis.set_major_formatter(
    mtick.FuncFormatter(lambda y, _: f"{y/1e6:.1f}M")
)

ax.legend(["Sales", "Profit"], frameon=False)
ax.tick_params(axis="x", rotation=30)

sns.despine()
plt.tight_layout()
plt.show()


# ----------------------------------------------------------
# 5. MONTHLY SALES HEATMAP
# ----------------------------------------------------------

df["order_year"] = df["order_date"].dt.year
df["order_month_number"] = df["order_date"].dt.month

monthly_heatmap = df.pivot_table(
    index="order_year",
    columns="order_month_number",
    values="sales",
    aggfunc="sum"
)

fig, ax = plt.subplots(figsize=(13, 6))

sns.heatmap(
    monthly_heatmap,
    cmap="YlGnBu",
    annot=True,
    fmt=".0f",
    linewidths=0.5,
    cbar_kws={"label": "Sales"},
    ax=ax
)

ax.set_title("Monthly Sales Performance by Year", pad=18)
ax.set_xlabel("Month")
ax.set_ylabel("Year")

month_names = [
    "Jan", "Feb", "Mar", "Apr", "May", "Jun",
    "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
]

ax.set_xticklabels(month_names)

plt.tight_layout()
plt.show()


# ----------------------------------------------------------
# 6. PROFIT BY CATEGORY
# ----------------------------------------------------------

profit_by_category = (
    df.groupby("category")["profit"]
    .sum()
    .sort_values()
)

fig, ax = plt.subplots(figsize=(9, 5))

bars = ax.barh(
    profit_by_category.index,
    profit_by_category.values,
    color=secondary_color
)

ax.set_title("Profit Contribution by Category", pad=18)
ax.set_xlabel("Total Profit")

ax.xaxis.set_major_formatter(
    mtick.FuncFormatter(lambda x, _: f"{x/1e6:.2f}M")
)

ax.bar_label(
    bars,
    labels=[f"{x/1e6:.2f}M" for x in profit_by_category.values],
    padding=5
)

sns.despine()
plt.tight_layout()
plt.show()

print("\nAll visualizations generated successfully!")


# ==========================================================
# STEP 6: FINAL SUMMARY
# ==========================================================

print("\n========== GLOBAL SUPERSTORE SUMMARY ==========")

print("Total Sales:", round(df['sales'].sum(), 2))

print("Total Profit:", round(df['profit'].sum(), 2))

print("Total Orders:", df['order_id'].nunique())

print("Total Customers:", df['customer_id'].nunique())

print("Total Products:", df['product_id'].nunique())

print("Total Quantity Sold:", df['quantity'].sum())

print("\nAnalysis Completed Successfully!")
