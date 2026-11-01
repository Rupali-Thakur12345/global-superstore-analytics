# 📊 Global Superstore Data Analytics Project

## 📌 Project Overview

The Global Superstore Data Analytics Project focuses on analyzing retail sales data to understand business performance, customer behaviour, product performance, and profitability across different regions and markets.

The project uses Python for data cleaning, Exploratory Data Analysis (EDA), and data visualization to identify meaningful business insights and support data-driven decision-making.

The dataset contains 51,290 records of global retail transactions, covering product categories, customers, sales, profits, shipping details, and geographical information.

## 🎯 Project Objectives

- Analyze overall sales and profit performance.
- Understand customer purchasing behaviour and segment performance.
- Identify high-value customers.
- Analyze category-wise and sub-category-wise sales.
- Evaluate regional and market-wise performance.
- Identify monthly sales trends.
- Analyze shipping methods and their impact on profitability.
- Identify business opportunities and areas requiring improvement.

## 🛠️ Tools & Technologies

| Tool       | Purpose                        |
| ---------- | ------------------------------ |
| Python     | Data analysis                  |
| Pandas     | Data cleaning and manipulation |
| Matplotlib | Data visualization             |
| Seaborn    | Statistical visualization      |

## 📂 Dataset Information

- **Dataset Name:** Global Superstore
- **Total Records:** 51,290
- **Original Columns:** 27
- **Final Columns:** 26
- **Missing Values:** 0
- **Duplicate Rows:** 0
- **Memory Usage After Cleaning:** Approximately 10.2 MB

### Important Columns

| Column        | Description                |
| ------------- | -------------------------- |
| order_id      | Unique order identifier    |
| order_date    | Order placement date       |
| ship_date     | Shipping date              |
| customer_id   | Unique customer identifier |
| customer_name | Customer name              |
| category      | Product category           |
| sub_category  | Product sub-category       |
| sales         | Sales amount               |
| profit        | Profit earned              |
| quantity      | Quantity sold              |
| discount      | Discount applied           |
| region        | Geographical region        |
| market        | Sales market               |
| segment       | Customer segment           |
| ship_mode     | Shipping method            |
| shipping_cost | Shipping cost              |

## 🔄 Project Workflow

### 1. Data Loading

- Imported the CSV dataset using Pandas.
- Inspected the dataset using `head()`, `info()`, and `describe()`.
- Reviewed column names and data types.

### 2. Data Cleaning

- Standardized column names.
- Checked missing values and duplicate records.
- Removed the unnecessary `row_id` column.
- Converted order and shipping dates into datetime format.
- Reviewed and optimized data types.

### 3. Exploratory Data Analysis (EDA)

Performed analysis on:

- Overall sales and profit.
- Category-wise and sub-category-wise sales.
- Customer purchasing patterns.
- Customer segment performance.
- Top 10 customers by sales.
- Region-wise and market-wise performance.
- Monthly sales trends.
- Shipping mode analysis.

### 4. Data Visualization

Created visualizations to understand business performance:

- Sales by Category.
- Monthly Sales Trend.
- Top 10 Sub-Categories by Sales.
- Regional Sales vs. Profit.
- Monthly Sales Heatmap.
- Profit by Category.
- KPI Summary Dashboard.

## 📈 Key Business Insights

### 1. Overall Business Performance

| KPI                 |       Result |
| ------------------- | -----------: |
| Total Sales         |   12,642,905 |
| Total Profit        | 1,467,457.29 |
| Total Orders        |       25,035 |
| Total Customers     |        4,873 |
| Total Products      |       10,292 |
| Total Quantity Sold |      178,312 |

### 2. Sales by Category

| Category        | Total Sales |
| --------------- | ----------: |
| Technology      |   4,744,691 |
| Furniture       |   4,110,884 |
| Office Supplies |   3,787,330 |

**Key Insight:** Technology generated the highest sales, followed by Furniture and Office Supplies.

### 3. Profit by Category

| Category        | Total Profit |
| --------------- | -----------: |
| Technology      |   663,778.73 |
| Office Supplies |   518,473.83 |
| Furniture       |   285,204.72 |

**Key Insight:** Technology contributed the highest profit, while Furniture recorded the lowest profit among the three categories.

### 4. Customer Behaviour Analysis

Customer behaviour was analyzed using customer sales, customer segments, and top-performing customers.

**A. Customer Segment Analysis**

| Customer Segment | Total Sales |
| ---------------- | ----------: |
| Consumer         |   6,508,141 |
| Corporate        |   3,824,808 |
| Home Office      |   2,309,956 |

**Key Insights:**

- Consumer customers contributed the highest sales.
- Corporate customers were the second-largest sales contributor.
- Home Office customers generated the lowest sales among the three segments.
- The Consumer segment represents an important customer group for targeted marketing and retention strategies.

**B. Top 10 Customers by Sales**

| Customer Name      | Total Sales |
| ------------------ | ----------: |
| Tom Ashbrook       |      40,489 |
| Tamara Chand       |      37,453 |
| Greg Tran          |      35,552 |
| Christopher Conant |      35,187 |
| Sean Miller        |      35,170 |
| Bart Watters       |      32,315 |
| Natalie Fritzler   |      31,778 |
| Fred Hopkins       |      30,404 |
| Jane Waco          |      30,288 |
| Hunter Lopez       |      30,246 |

**Key Insights:**

- Tom Ashbrook was the highest sales-generating customer.
- Tamara Chand and Greg Tran were the second and third highest contributors.
- The top customers represent valuable accounts for customer retention initiatives.

**C. Customer Behaviour Recommendations**

- Introduce loyalty programs to retain high-value customers.
- Develop targeted offers for Consumer customers.
- Analyze purchasing frequency to identify repeat customers.
- Study customer preferences to recommend relevant products.
- Create personalized marketing campaigns for different customer segments.

### 5. Sub-Category Analysis

**Top 5 Sub-Categories by Sales**

| Sub-Category | Total Sales |
| ------------ | ----------: |
| Phones       |   1,706,874 |
| Copiers      |   1,509,439 |
| Chairs       |   1,501,682 |
| Bookcases    |   1,466,559 |
| Storage      |   1,127,124 |

**Key Insight:** Phones generated the highest sales among all sub-categories.

### 6. Profitability Analysis

- Copiers generated the highest sub-category profit of 258,567.55.
- Phones generated a profit of 216,717.01.
- Tables recorded a negative profit of -64,083.39.
- Furniture profitability requires further investigation, particularly in the Tables sub-category.

### 7. Regional Performance

| Region         |     Sales |     Profit |
| -------------- | --------: | ---------: |
| Central        | 2,822,399 | 311,403.98 |
| South          | 1,600,960 | 140,355.77 |
| North          | 1,248,192 | 194,597.95 |
| Oceania        | 1,100,207 | 120,089.11 |
| Southeast Asia |   884,438 |  17,852.33 |

**Key Insight:** Central recorded the highest sales and profit among the regions.

### 8. Market Performance

| Market |     Sales |     Profit |
| ------ | --------: | ---------: |
| APAC   | 3,585,833 | 436,000.05 |
| EU     | 2,938,139 | 372,829.74 |
| US     | 2,297,354 | 286,397.02 |
| LATAM  | 2,164,687 | 221,643.49 |
| EMEA   |   806,184 |  43,897.97 |

**Key Insight:** APAC was the highest-performing market in terms of both sales and profit.

### 9. Shipping Mode Analysis

| Shipping Mode  |     Sales |     Profit |
| -------------- | --------: | ---------: |
| Standard Class | 7,578,889 | 890,596.02 |
| Second Class   | 2,565,747 | 292,583.53 |
| First Class    | 1,831,067 | 208,104.68 |
| Same Day       |   667,202 |  76,173.07 |

**Key Insight:** Standard Class generated the highest sales and profit among all shipping methods.

## 💡 Business Recommendations

1. **Customer Retention:** Introduce loyalty programs and personalized offers for high-value customers.
2. **Customer Segmentation:** Focus marketing strategies on the Consumer segment while exploring growth opportunities in Corporate and Home Office.
3. **Product Strategy:** Prioritize high-performing products such as Phones and Copiers.
4. **Profitability Improvement:** Investigate the negative profit in Tables and review discount and pricing strategies.
5. **Regional Growth:** Study successful markets such as APAC and regions such as Central to identify opportunities for other markets.
6. **Shipping Optimization:** Evaluate shipping costs and operational efficiency across different shipping methods.

## 👩‍💻 Author

**Rupali Thakur**

Aspiring Data Analyst

**Skills:** Python | Pandas | SQL | Advanced Excel | Power BI | Data Visualization

- GitHub: [Rupali-Thakur12345](https://github.com/Rupali-Thakur12345)
- LinkedIn: [LinkedIn Profile](https://www.linkedin.com/in/rupali-thakur-a55aa9281/)

---

*This project demonstrates practical experience in Python-based data analysis, customer behaviour analysis, exploratory data analysis, and business insights generation.*
