# Sales Performance Dashboard — Excel
> Cleaned and analyzed Q1-Q2 2025 sales data in Excel, building an interactive dashboard with Pivot Tables, slicers, and KPI cards to uncover revenue trends, top products, and performance insights across cities, channels, and sales reps.

---

## ⚙️ Project Type Flags

- [ ] Exploratory Data Analysis (EDA)
- [ ] SQL Analysis / Querying
- [x] Dashboard / Data Visualization
- [ ] Data Pipeline / ETL
- [ ] Predictive Modelling / Machine Learning
- [x] Data Cleaning / Wrangling
- [ ] End-to-End (multiple of the above)
- [ ] Other:

---

## Table of Contents
1. [Project Overview](#1-project-overview)
2. [Objectives](#2-objectives)
3. [Project Scope & Tools](#3-project-scope--tools)
4. [Repository Structure](#4-repository-structure)
5. [Data Workflow](#5-data-workflow)
6. [Data Model & Schema](#6-data-model--schema)
7. [Analysis & Metrics](#7-analysis--metrics)
8. [Key Insights](#8-key-insights)
9. [Recommendations](#9-recommendations)
10. [Deliverables](#10-deliverables)
11. [Author](#11-author)

---

## 1. Project Overview

**Context:** A retail sales business needed a clear view of its sales and profitability performance across Q1-Q2 2025. The business had raw transaction data spread across multiple records but no structured summary to track revenue, profit, top performers, or sales trends.

**Problem Statement:** Without a visual summary, the business could not easily identify which sales reps were performing best, which cities generated the most revenue, which products were top sellers, or how sales trended across months and days.

**Approach:** Cleaned and structured raw sales data in Excel, building a Sales Data sheet and a lookup table. Used Pivot Tables to summarize performance across multiple dimensions including city, channel, category, sales rep, and time period. Designed an interactive dashboard with KPI cards, slicers, and multiple chart types to enable dynamic exploration of the data.

**Outcome:** Delivered a fully interactive Excel Sales Dashboard revealing Total Revenue, Total Profit, Total Orders, Profit Margin, and Average Order Value - alongside visual breakdowns of monthly trends, order by city, quantity sold by product, revenue by category, profit by sales rep, order by channel, and order by day of week.

---

## 2. Objectives

- **Primary Objective:** Build an interactive Excel sales dashboard to visualize Q1-Q2 2025 sales performance across revenue, profit, products, cities, channels, and sales representatives.
- **Secondary Objective 1:** Identify top-performing products and categories by total revenue and order volume.
- **Secondary Objective 2:** Compare sales performance across cities, sale channels, and individual sales representatives to identify strengths and gaps.
- **Secondary Objective 3:** Analyze monthly and daily sales trends to identify peak and low performance periods across the Q1-Q2 period.
- **Secondary Objective 4:** Calculate and display key business metrics including Total Revenue, Total Profit, Profit Margin, Total Orders, and Average Order Value using Pivot Tables and Excel formulas.

> 💡 Every analysis decision in this project traces back to one of these objectives.

---

## 3. Project Scope & Tools

### Scope

| Dimension | Details |
|------------|---------|
| **In Scope** | Q1-Q2 2025 sales transaction records covering revenue, profit, orders, cities, channels, categories, products, and sales representatives |
| **Out of Scope** | Customer demographic data, inventory levels, and post-sale returns - these were not available in the dataset |
| **Time Period** | January 2025 - June 2025 (Q1-Q2) |
| **Granularity** | Row-level transaction data (one row per sales order) |

### Tools & Technologies

| Category | Tool(s) Used |
|----------|--------------|
| Data Cleaning | Microsoft Excel (deduplication, formatting, standardization) |
| Data Analysis | Microsoft Excel (Pivot Tables, VLOOKUP, COUNTIF, SUM, IF formulas) |
| Data Visualization | Microsoft Excel (bar charts, line charts, donut charts, KPI cards) |
| Dashboard Design | Microsoft Excel (interactive dashboard with slicers, KPI cards, dynamic charts) |
| Documentation | Microsoft Word, GitHub |

---

## 4. Repository Structure

```
Sales-Performance-Dashboard/
|
├── data/
|   └── raw/              # Original, unmodified sales dataset
|
├── docs/                 # Data dictionary and project notes
|
├── reports/              # Written summary report
|
├── visuals/              # Dashboard screenshots and charts
|
├── README.md             # You are here
└── project_metadata.yml  # Project metadata

```

---

## 5. Data Workflow

1. **Source:** One Excel workbook containing raw sales transaction data for Q1–Q2 2025, capturing order dates, product details, quantities, cities, sale channels, sales representatives, revenue, and profit figures.

2. **Ingestion:** Dataset loaded and opened directly in Microsoft Excel for cleaning and analysis.

3. **Cleaning:** The following data quality steps were performed:
   - Removed duplicate transaction records
   - Standardized date formatting across the order date field
   - Ensured consistent naming conventions for cities, channels, categories, and sales representatives
   - Verified numerical fields (revenue, profit, quantity) for accuracy and completeness

4. **Transformation:** Built a structured Sales Data sheet and a lookup table to enable Pivot Table analysis. Used VLOOKUP to join related data across sheets where needed.

5. **Analysis:** Built multiple Pivot Tables to summarize performance across cities, channels, categories, products, sales reps, months, and days. Calculated KPI metrics including Total Revenue, Total Profit, Profit Margin, Total Orders, and Average Order Value using Excel formulas.

6. **Visualization:** Designed an interactive dashboard with KPI cards, slicers for filtering by city, channel, and time period, and multiple chart types including bar charts, line charts, donut charts, and category breakdowns.

7. **Output:** Interactive Excel dashboard, written summary report (Word document), dashboard screenshots uploaded to GitHub, and full project documentation in this repository.

---

## 6. Data Model & Schema

### Dataset Description
The dataset contains Q1-Q2 2025 sales transaction records for a retail sales business.

### Data Dictionary

| Field Name | Data Type | Description | Example Value |
|------------|-----------|-------------|---------------|
| `Date` | Date | Date of the sales transaction | 2025-01-19 |
| `Region` | Text | Geographic region of the sale | West |
| `City` | Text | City where the sale occurred | Port Harcourt |
| `Customer Type` | Text | Type of customer (Retail or Corporate) | Retail |
| `Channel` | Text | Sales channel (Online or Store) | Online |
| `Product` | Text | Name of the product sold | Office Chair Pro |
| `Category` | Text | Product category | Furniture |
| `Unit Price` | Integer | Selling price per unit | 85,000 |
| `Quantity` | Integer | Number of units sold | 1 |
|`Sales Rep` | Text | Name of the sales representative | Chidinma |
| `Cost Price` | Integer | Cost price per unit | 68,000 |
| `Revenue` | Integer | Total revenue from the transaction | 85,000 |
| `COGS` | Integer | Cost of Goods Sold | 68,000 |
| `Profit` | Integer | Revenue minus COGS | 17,000 |
| `Month` | Text | Month of the transaction | January |
| `Day` | Text | Day of the week | Tuesday |

> *Date range:* January 2025 - June 2025 (Q1-Q2)
> *Regions covered:* North, South, East, West
> *Cities covered:* Lagos, Kano, Abuja, Port Harcourt
> *Sales Representatives:* Chidinma, David, Musa, Peter, Grace, Aisha
> *Categories:* Electronics, Furniture, Home Appliances
> 
---

## 7. Analysis & Metrics

### Analytical Approach

This project used a *dashboard-driven analysis* approach - structuring and summarizing Q1-Q2 2025 sales data using Pivot Tables and Excel formulas to build an interactive dashboard that enables business stakeholders to explore performance across multiple dimensions dynamically.

### Key Metrics Defined

| Metric | Plain-Language Definition | Why It Matters |
|--------|--------------------------|----------------|
| Total Revenue | Sum of all revenue across transactions | Measures overall business income generated |
| Total Profit | Sum of all profit across transactions | Shows actual earnings after cost of goods sold |
| Profit Margin | Total Profit divided by Total Revenue multiplied by 100 | Measures how efficiently revenue converts to profit |
| Total Orders | Count of all sales transactions | Measures overall sales activity volume |
| Average Order Value | Total Revenue divided by Total Orders | Shows the average value of each transaction |
| Revenue by City | Total revenue grouped by city | Identifies strongest and weakest performing locations |
| Revenue by Category | Total revenue grouped by product category | Shows which product categories drive the most income |
| Profit by Sales Rep | Total profit grouped by sales representative | Identifies top and low performing sales team members |
| Orders by Channel | Total orders split by Online vs Store | Compares performance across sales channels |
| Orders by Day | Total orders grouped by day of week | Reveals which days record the highest sales activity |
| Monthly Trend | Total revenue and profit grouped by month | Shows performance trajectory across Q1-Q2 |

### Methods Used

- Data cleaning - deduplication, date formatting, naming standardization
- Pivot Tables - summarizing revenue, profit, orders across cities, channels, categories, reps, months and days
- VLOOKUP - joining related data across sheets
- KPI formulas - calculating Total Revenue, Total Profit, Profit Margin, Total Orders, Average Order Value
- Data visualization - KPI cards, bar charts, line charts, donut charts, slicers for dynamic filtering

---

## 8. Key Insights

1. **Total Revenue is ₦2,328,370,000 with a 20% Profit Margin** - generating Total Profit of ₦465,674,000 across 2,098 total orders, with an Average Order Value of ₦1,109,805.

2. **Port Harcourt is the top performing city** with the highest order volume at 575,480,000 - significantly ahead of Lagos, Kano, and Abuja.

3. **February recorded the highest monthly revenue** peaking at 633,440,000 - with a sharp drop in May to 11,100,000, indicating a heavily front-loaded sales period.

4. **Laptop A13 is the top product** with revenue of ₦526,080,000 - making it the single most valuable product in the catalogue.

5. **Electronics is the top performing category** generating ₦1,323,880,000 in revenue - significantly ahead of Furniture and Home Appliances.

6. **Sofa Classic leads in quantity sold** followed by Air Conditioner X2 and Blender B10 - showing strong volume in furniture and home appliances despite Electronics leading in revenue.

7. **Online channel dominates over Store** - Online transactions account for the larger share of orders across the Q1–Q2 period.

8. **Sales performance varies significantly across representatives** - with clear top and bottom performers visible across Aisha, Chidinma, David, Grace, Musa, and Peter.

---

## 9. Recommendations

| Priority | Recommendation | Based On | Suggested Owner |
|----------|---------------|----------|-----------------|
| High | Investigate the sharp revenue drop in May (₦11,100,000 vs February peak of ₦633,440,000) and implement strategies to sustain sales momentum beyond Q1 | Insight 3 - February peaked at ₦633,440,000 with a sharp May decline | Sales / Marketing team |
| High | Prioritize stock availability for Laptop A13 year-round given its status as the top revenue-generating product at ₦526,080,000 | Insight 4 - Laptop A13 is the single most valuable product | Inventory / Procurement team |
| High | Develop targeted growth strategies for Port Harcourt to sustain its leading position, while identifying what is holding back Abuja at ₦529,140,000 | Insight 2 - Port Harcourt leads at ₦575,480,000 while Abuja trails | Regional Sales team |
| Medium | Expand Electronics category offerings given its dominant revenue contribution of ₦1,323,880,000 - nearly double Furniture and Home Appliances combined | Insight 5 - Electronics generates the highest category revenue | Product / Procurement team |
| Medium | Review underperforming sales representatives and provide targeted coaching or incentives to close the performance gap with top performers | Insight 8 - Significant variation across sales rep performance | Sales Management team |
| Medium | Develop strategies to grow the Store channel to reduce over-reliance on Online sales | Insight 7 - Online channel dominates over Store | Sales / Marketing team |
| Low | Investigate high-volume but lower-revenue products like Sofa Classic to determine if pricing adjustments could improve profit margins | Insight 6 - Sofa Classic leads in quantity but Electronics leads in revenue | Pricing / Product team |

---

## 10. Deliverables

| Deliverable | Description | Location |
|-------------|-------------|----------|
| Excel Sales Dashboard | Interactive dashboard with KPI cards, slicers, and charts covering revenue, profit, products, cities, channels, sales reps, and trends | visuals/Sales_Performance_Dashboard.png |
| Raw Dataset | Original Excel file containing Sales Data sheet with all transaction records | data/raw/sales_data.xlsx |
| Summary Report | Written Word document summarizing findings, insights, and recommendations | reports/Sales_Performance_Dashboard_Summary_Report.docx |

---

## 11. Author

* Vivian Okwara *
Data Analyst | Lagos, Nigeria

- 🔗 https://linkedin.com/in/okwara-vivian
- 🌐 https://Vivian-Portfolio.github.io
- 📧 laviv030@gmail.com

---

Last updated: June 2026




