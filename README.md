# Excel-project
# 🛒 E-Commerce Customer Insights & Churn Analysis

![Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![Python](https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white)
![Data Analysis](https://img.shields.io/badge/Data%20Analysis-FF6F00?style=for-the-badge&logo=databricks&logoColor=white)
![Status](https://img.shields.io/badge/Status-Complete-brightgreen?style=for-the-badge)

---

## 📌 Project Overview

This project performs a comprehensive analysis of an e-commerce platform's customer behavior, subscription patterns, and churn risk. Using a dataset of **2,000 customer transactions** spanning 6 countries and 5 product categories, the analysis uncovers actionable insights on revenue performance, customer retention, and demographic trends.

The project is structured as an end-to-end Excel analytics workflow — from raw data cleaning through KPI computation, pivot table summarization, chart visualization, and an executive-level dashboard.

---

## 🎯 Objectives

- Identify customers at risk of churning based on cancellation behavior and subscription status
- Calculate key business KPIs: revenue, average order value, churn rate, and purchase frequency
- Segment customers by age group, gender, country, and product category
- Visualize performance trends through pivot charts and an interactive dashboard
- Provide data-driven recommendations for improving customer retention

---

## 📁 File Descriptions

| File | Type | Description |
|------|------|-------------|
| `E_Commerce_Customer_Insights_and_Churn_Dataset.csv` | Raw Data | Source dataset with 2,000 customer order records and 17 features |
| `Excel_Project.xlsx` | Analysis Workbook | Full Excel project including cleaned data, KPIs, pivot tables, charts, and dashboard |

---

## 🗂️ Dataset Description

- **Source File:** `E_Commerce_Customer_Insights_and_Churn_Dataset.csv`
- **Records:** 2,000 rows (one per order/customer)
- **Features:** 17 columns
- **Geography:** Canada, USA, UK, Germany, India, Pakistan
- **Categories:** Sports, Clothing, Electronics, Home, Beauty
- **Date Range:** Spans multiple years of signup and purchase activity

---

## 🔢 Column / Feature Explanation

| Column | Data Type | Description |
|--------|-----------|-------------|
| `order_id` | String | Unique identifier for each order (e.g., `ORD5000`) |
| `customer_id` | String | Unique identifier for each customer (e.g., `CUST1000`) |
| `age` | Integer | Customer's age (range: 18–69) |
| `product_id` | String | Unique identifier for the purchased product (e.g., `PROD200`) |
| `country` | String | Customer's country of residence (Canada, USA, UK, Germany, India, Pakistan) |
| `signup_date` | Date | Date the customer registered on the platform |
| `last_purchase_date` | Date | Date of the customer's most recent purchase |
| `cancellations_count` | Integer | Total number of subscription cancellations by the customer (range: 0–5) |
| `subscription_status` | String | Current subscription state: `active`, `cancelled`, or `paused` |
| `order_date` | Date | Date the specific order was placed |
| `unit_price` | Float | Price per unit of the ordered product (range: $2.85–$1,991.63) |
| `quantity` | Integer | Number of units purchased in the order |
| `purchase_frequency` | Integer | How frequently the customer makes purchases (range: 1–49) |
| `preferred_category` | String | The product category the customer prefers (Sports, Clothing, Electronics, Home, Beauty) |
| `product_name` | String | Name of the specific product ordered (e.g., Football, Refrigerator, Hoodie) |
| `category` | String | Actual category of the product ordered |
| `gender` | String | Customer's gender: `Male`, `Female`, or `Other` |

> **💡 Churn Indicator:** Customers with `subscription_status = 'cancelled'` and high `cancellations_count` (4–5) are treated as high churn-risk profiles in the analysis.

---

## 📊 Excel Workbook Sheet Descriptions

The `Excel_Project.xlsx` file contains **6 sheets**, each serving a specific purpose in the analytics pipeline:

| Sheet | Purpose |
|-------|---------|
| **Raw Data** | Original imported dataset — not modified |
| **Data** | Cleaned and transformed data with derived columns (e.g., gross revenue, age groups, churn risk flags) |
| **KPI Calculation** | Executive summary table with 18 calculated business KPIs |
| **Pivot Table** | Multi-dimensional summaries by category, country, gender, age group, and subscription status |
| **Pivot Charts** | Visual representations of the pivot table data |
| **Dashboard** | Interactive executive dashboard combining all key visuals and KPIs |

---

## 📈 Key KPIs (from KPI Calculation Sheet)

| KPI | Value |
|-----|-------|
| Total Gross Revenue | **$2,051,690.65** |
| Total Orders | **2,000** |
| Unique Customers | **2,000** |
| Average Order Value (AOV) | **$1,025.85** |
| Active Customer Rate | **60.2%** |
| Churn Rate (Cancelled) | **24.65%** |
| Paused Rate | **15.15%** |
| High Churn Risk Count | **388 customers** |
| Top Revenue Category | **Clothing — $439,803.15** |
| Top Revenue Country | **Germany — $368,249.31** |
| Avg Cancellations / Customer | **2.44** |
| Total Cancellations | **4,883** |

---

## 🛠️ Technologies Used

| Tool | Purpose |
|------|---------|
| **Microsoft Excel** | Data cleaning, KPI calculation, pivot tables, charts, dashboard |
| **Python (pandas)** | Data inspection and validation |
| **CSV** | Raw data storage and transfer format |
| **Markdown** | Project documentation |

---

## ⚙️ Installation Instructions

### Prerequisites

Make sure you have the following installed:

- **Microsoft Excel** 2016 or later (for full pivot chart and dashboard support)
  - *Alternatively:* LibreOffice Calc (limited chart support) or Google Sheets
- **Python 3.8+** (optional, for data exploration via script)

### Setup Steps

1. **Clone or download this repository:**

   ```bash
   git clone https://github.com/your-username/ecommerce-churn-analysis.git
   cd ecommerce-churn-analysis
   ```

2. **Install Python dependencies** (optional):

   ```bash
   pip install pandas openpyxl
   ```

3. **Open the Excel workbook:**

   ```
   Double-click Excel_Project.xlsx
   ```

4. **Enable macros/content** if prompted by Excel security settings.

---

## 🚀 Usage Guide

### Using the Excel Workbook

1. Open `Excel_Project.xlsx`
2. Start on the **Dashboard** sheet for a high-level executive summary
3. Navigate to **Pivot Table** to filter and explore data by category, country, or gender
4. Visit **KPI Calculation** for a full breakdown of business metrics
5. Use **Pivot Charts** for visual trend analysis

### Using the Raw CSV (Python)

```python
import pandas as pd

# Load the dataset
df = pd.read_csv("E_Commerce_Customer_Insights_and_Churn_Dataset.csv")

# Quick overview
print(df.shape)         # (2000, 17)
print(df.head())
print(df.describe())

# Check subscription distribution
print(df['subscription_status'].value_counts())

# Filter high churn-risk customers
high_risk = df[df['cancellations_count'] >= 4]
print(f"High churn risk customers: {len(high_risk)}")
```

---

## 🧹 Data Preprocessing Steps

The following transformations were applied in the **Data** sheet:

1. **Date Standardization** — Converted Excel serial date numbers to readable date format (`DD/MM/YYYY`) for `signup_date`, `last_purchase_date`, and `order_date`

2. **Gross Revenue Calculation** — Derived a new column:
   ```
   gross_revenue = unit_price × quantity
   ```

3. **Age Group Segmentation** — Bucketed the `age` column into 5 generational bands:

   | Age Range | Label |
   |-----------|-------|
   | 18–24 | Gen Z |
   | 25–34 | Millennials |
   | 35–49 | Gen X |
   | 50–64 | Boomers |
   | 65+ | Seniors |

4. **Churn Risk Flagging** — Assigned a `churn_risk` label based on `subscription_status` and `cancellations_count`:
   - `High` → cancelled status or cancellations ≥ 4
   - `Medium` → paused status or cancellations 2–3
   - `Low` → active status and cancellations ≤ 1

5. **Preference Alignment Check** — Boolean flag indicating whether `preferred_category` matches the actual `category` of the order placed

6. **Duplicate & Null Checks** — Verified there are no duplicate `order_id` values and no missing critical fields

---

## 📋 Example Outputs / Results

### Subscription Status Breakdown

| Status | Customers | % Share |
|--------|-----------|---------|
| Active | 1,204 | 60.2% |
| Cancelled | 493 | 24.65% |
| Paused | 303 | 15.15% |

### Revenue by Product Category

| Category | Revenue | Orders | Avg Order Value |
|----------|---------|--------|-----------------|
| Clothing | $439,803 | 426 | $1,032.40 |
| Electronics | $411,629 | 414 | $994.27 |
| Beauty | $405,851 | 396 | $1,024.88 |
| Sports | $402,321 | 390 | $1,031.59 |
| Home | $392,087 | 374 | $1,048.36 |

### Revenue by Country

| Country | Revenue | Orders |
|---------|---------|--------|
| Germany | $368,249 | 360 |
| UK | $353,312 | 350 |
| Pakistan | $343,163 | 332 |
| India | $336,792 | 324 |
| USA | $328,657 | 319 |
| Canada | $321,518 | 315 |

### Revenue by Age Group

| Age Group | Customers | Revenue | Avg Order Value |
|-----------|-----------|---------|-----------------|
| Gen Z (18–24) | 258 | $263,546 | $1,021.50 |
| Millennials (25–34) | 345 | $361,799 | $1,048.69 |
| Gen X (35–49) | 584 | $598,548 | $1,024.91 |
| Boomers (50–64) | 632 | $633,398 | $1,002.21 |
| Seniors (65+) | 181 | $194,398 | $1,074.02 |

---

## 🗃️ Folder Structure

```
ecommerce-churn-analysis/
│
├── 📄 README.md                                          # Project documentation (this file)
├── 📊 Excel_Project.xlsx                                 # Full Excel analytics workbook
└── 📋 E_Commerce_Customer_Insights_and_Churn_Dataset.csv # Raw source dataset
```

---

## 🔮 Future Improvements

- [ ] **Machine Learning Model** — Build a churn prediction classifier (Logistic Regression / Random Forest) using Python's `scikit-learn`
- [ ] **Power BI / Tableau Dashboard** — Migrate the Excel dashboard to an interactive BI tool for real-time filtering
- [ ] **Time-Series Analysis** — Analyze revenue and churn trends over time using `signup_date` and `order_date`
- [ ] **Customer Lifetime Value (CLV)** — Calculate and segment customers by projected long-term value
- [ ] **Cohort Analysis** — Group customers by signup month to track retention curves over time
- [ ] **Automated Reporting** — Schedule Python scripts to auto-generate updated KPI summaries from refreshed CSV exports
- [ ] **Geospatial Visualization** — Map revenue and churn distribution by country using `folium` or `plotly`

---


---

## 👤 Author Information

*Mohammed Taqi* 


> ⭐ If you found this project helpful, consider giving it a star on GitHub!
