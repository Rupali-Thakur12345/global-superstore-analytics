import pandas as pd

file_path = r"C:\Users\user\Downloads\archive (1)\Global_Superstore2.xlsx"
df =pd.read_excel(file_path,engine ='openpyxl')
# # Preview first few row
print(df.head())
# # show data types and missing values
print(df.info())
print(df.isnull().sum)

#STEP 2: Data Cleaning & Preparation
# Step 2.1 - Remove duplicates
print("before removing duplicates:",df.shape)
df.drop_duplicates(inplace=True)
print("after removing duplicates:", df.shape)

# Step 2.2 Check data types
print(df.dtypes)

# Step 2.3 - Convert 'order date' and 'ship date ' to database

df['Order Date']= pd.to_datetime(df ['Order Date'])
df['Ship Date'] = pd.to_datetime(df['Ship Date'])

# Srep 2.4 - Drop 'row id' column (not needed for analysis)
df.drop(columns=['Row ID'], inplace= True)

# Step 2.5 - confire final structur
print(df.info())

#STEP 3 :Exploratory data analysis(EDA)
# 3.1
#  Toatal sales by category     

sales_by_category =df.groupby("Category")['Sales'].sum()

print("\ total sales by category:")
print(sales_by_category)

# Total sales by sub_category
sales_by_sub_category =df.groupby("Sub-Category")['Sales'].sum()

print("\ total sales by sub-category:")
print(sales_by_sub_category)


# 3.2 Most profitable categories/sub -categories
profit_by_category =df.groupby("Category")['Profit'].sum()

print("\ total profit by category:")
print(profit_by_category)
print(df.groupby("Category"))
# profit by sub category
profit_by_sub_category =df.groupby("Sub-Category")['Profit'].sum()

print("\ total profit by sub-category:")
print(profit_by_sub_category)

#3.3 Region - wise sales and profit

region_summary =df.groupby("Region")[["Sales","Profit"]].sum().sort_values(by="Sales",
ascending=False)
print(region_summary)

# 3.4 Monthly sales trend
# add a new cplumn for month -year
df["Month"] = df["Order Date"].dt.to_period("M")
# monthly sales trend
monthly_sales = df.groupby("Month")["Sales"].sum()

print(monthly_sales.tail(12))                 # last 12 month

print(df.columns)
print(type(df))
grouped =df.groupby('Category')['Sales']
print(type(grouped))
print(grouped.sum())

# Data Visualization with matplotlib/seaborn

# 1. Bar Plot - Total Sales by category

import matplotlib.pyplot as plt
import seaborn as sns

# group data
sales_by_category = df.groupby('Category')['Sales'].sum().sort_values(ascending=False)

# plot
plt.figure(figsize=(8,5))
sns.barplot(x=sales_by_category.index,y=sales_by_category
.values,palette='pastel')
plt.title("Total Sales by Category",fontsize=14)
plt.ylabel("Sales")
plt.xlabel("Category")
plt.tight_layout()
plt.show()

#2. Pie Chart-profit by region

profit_by_region = df.groupby('Region')['Profit'].sum()

#plot
plt.figure(figsize=(8,5))
plt.pie(profit_by_region,
labels=profit_by_region.index, autopct='%1.1f%%',startangle=90,
colors=sns.color_palette("pastel"))
plt.title("Profit Distribution by Region")
plt.axis('equal')
plt.show()

# Line Chart - Monthly Sales Trend

df['Order Month'] = df['Order Date'].dt.to_period('M')
# Group by month
monthly_sales = df.groupby('Order Month')['Sales'].sum().sort_index()
# plot
plt.figure(figsize=(12,5))
monthly_sales.plot(marker='o',color='blue')
plt.title("Monthly Sales Trend")
plt.xlabel('Order Month')

#Line chart - monthly sales trend

#EXTRACT MONTH-YEAR
df['Order Month']= df['Order Date'].dt.to_period('M')

# Group by month
monthly_sales = df.groupby('Order Month')['Sales'].sum().sort_index()
# plot
plt.figure(figsize=(12,5))
monthly_sales.plot(marker='o',color='blue')
plt.title('Monthly Sales Trend')
plt.xlabel('Order Month')
plt.ylabel("Sales")
plt.grid(True)
plt.tight_layout()
plt.show()

# horizontal bar chart - top 10 sub- categories by sales

top_subcats = df.groupby('Sub-Category')['Sales'].sum().sort_values(ascending= False).head(10)
#plot
plt.figure(figsize=(10,6))
sns.barplot(x=top_subcats.values,
y=top_subcats.index,palette='viridis')
plt.title("top 10 sub-categories by sales")
plt.xlabel("Sales")
plt.ylabel("Sub-Category")
plt.tight_layout()
plt.show()














