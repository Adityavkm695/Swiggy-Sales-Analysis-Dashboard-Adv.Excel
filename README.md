# 🍔 Swiggy Sales Performance Dashboard – Food Delivery Analytics

_Converting raw transactional logistics into an interactive, enterprise-grade Excel dashboard to support real-time sales, regional, and consumer-behavior decisions._

---

## 📌 Table of Contents
- <a href="#overview">Overview</a>
- <a href="#business-problem">Business Problem</a>
- <a href="#dataset">Dataset</a>
- <a href="#tools--technologies">Tools & Technologies</a>
- <a href="#project-structure">Project Structure</a>
- <a href="#data-cleaning--preparation">Data Cleaning & Preparation</a>
- <a href="#data-modeling--formula-engineering">Data Modeling & Formula Engineering</a>
- <a href="#kpis--key-findings">KPIs & Key Findings</a>
- <a href="#dashboard">Dashboard</a>
- <a href="#how-to-run-this-project">How to Run This Project</a>
- <a href="#final-recommendations">Final Recommendations</a>
- <a href="#author--contact">Author & Contact</a>

---
<h2><a class="anchor" id="overview"></a>Overview</h2>

This project transforms over **1.97 lakh raw Swiggy order transactions** spanning multiple Indian states into a single, interactive Excel dashboard. It combines structured data modeling, advanced formula engineering, and dynamic slicers to give leadership a real-time, zero-clutter view of sales, regional performance, and consumer behavior — without relying on any complex BI tool.

---
<h2><a class="anchor" id="business-problem"></a>Business Problem</h2>

Swiggy processes lakhs of orders daily across cities, restaurants, and food categories, but that data was scattered across disconnected spreadsheets. The business needed one consolidated view to:
- Track total sales, order volume, and average order value in real time
- Compare regional and city-wise performance
- Identify monthly, weekly, and daily sales trends
- Spot top-performing restaurants, cities, and food categories
- Enable leadership to filter and drill down instantly, without technical overhead

---
<h2><a class="anchor" id="dataset"></a>Dataset</h2>

- Raw Swiggy order-level transactional data (CSV format)
- **1,97,430 records** across State, City, Order Date, Restaurant, Dish, Food Type, Price, Rating, and Rating Count
- Structured and normalized into an indexed Excel table (`Swiggy_Data`) for dynamic scalability

---

<h2><a class="anchor" id="tools--technologies"></a>Tools & Technologies</h2>

- Microsoft Excel (Pivot Tables, Structured Tables, Advanced Formulas)
- Excel Slicers & Timelines (Interactive Filtering)
- Power Query (Data Cleaning & Transformation)
- Chart Visualizations (Line, Bar, Treemap, Geographic Map)
- GitHub

---
<h2><a class="anchor" id="project-structure"></a>Project Structure</h2>

```
swiggy-sales-performance-dashboard/
│
├── README.md
├── Swiggy_Sales_Dashboard_Executive_Brief.pdf
│
├── data/                        # Raw and cleaned transactional data
│   └── swiggy_raw_orders.csv
│
├── dashboard/                   # Final Excel dashboard file
│   └── swiggy_sales_dashboard.xlsx
│
├── images/                      # Dashboard screenshots
│   └── dashboard_preview.png
```

---
<h2><a class="anchor" id="data-cleaning--preparation"></a>Data Cleaning & Preparation</h2>

- Removed duplicate and blank rows from raw CSV logs
- Standardized inconsistent date formats and text-based numeric fields
- Fixed mismatched city/category naming conventions
- Normalized raw data into a structured, indexed table (`Swiggy_Data`) so all pivots and charts scale automatically as new data is appended

---
<h2><a class="anchor" id="data-modeling--formula-engineering"></a>Data Modeling & Formula Engineering</h2>

**Time-Series Formulas:**
- `=TEXT(Order_Date, "mmm")` → Extracts monthly trends
- `=TEXT(Order_Date, "ddd")` → Maps weekday vs. weekend performance
- `="Q" & ROUNDUP(MONTH(Order_Date)/3,0)` → Groups data into business quarters

**KPI Formulas:**
- Total Sales → `SUM(Price)`
- Total Orders → `COUNT(State/Order ID)`
- Average Order Value → `AVERAGE(Price)`
- Avg. Customer Rating → `AVERAGE(Rating)`

**Dynamic Chart-Slicer Bypass:**
Standard Excel charts don't auto-refresh with pivot slicers, so a mirror-range formula was used to keep every visual synced with filters in real time:
`=OFFSET(U3, 0, 0, COUNTA(U:U)-1, 2)`

---
<h2><a class="anchor" id="kpis--key-findings"></a>KPIs & Key Findings</h2>

| Metric | Value |
|---|---|
| Total Consolidated Sales | ₹5,30,12,506 |
| Total Order Fulfillments | 1,97,430 Orders |
| Average Order Value (AOV) | ₹268.51 |
| Platform Average Rating | 4.34 / 5.00 |
| Total Rating Volume | 55,91,574 Ratings |

**Key Insights:**
1. **Saturday Revenue Surge** – Saturdays consistently peak (up to ₹77,82,935/day), while Mondays/Tuesdays show order fatigue
2. **Seasonal Stability** – Revenue peaked in May (₹67.9L) and held steady through August, forming a summer demand baseline
3. **Dietary Split** – Vegetarian items drive 64% of revenue (₹3.38Cr) vs. Non-Veg at 36% (₹1.92Cr)
4. **Merchant Leaders** – KFC, McDonald's, and Pizza Hut dominate order density
5. **Geographic Leadership** – Bengaluru leads with ₹54.6L in sales, followed by Lucknow, Hyderabad, Mumbai, and New Delhi

---
<h2><a class="anchor" id="dashboard"></a>Dashboard</h2>

- Interactive Excel dashboard with:
  - KPI cards for Sales, Orders, AOV, and Rating
  - Monthly & daily trend line charts
  - Geographic distribution map
  - Top-10 food brand treemap
  - Quarterly performance matrix & regional city rankings
  - Slicers for Month, Food Type, State, and Category

![Swiggy Sales Dashboard](images/dashboard_preview.png)

---
<h2><a class="anchor" id="how-to-run-this-project"></a>How to Run This Project</h2>

1. Clone the repository:
```bash
git clone https://github.com/yourusername/swiggy-sales-performance-dashboard.git
```
2. Open the raw dataset in `data/swiggy_raw_orders.csv`
3. Open the Excel dashboard:
   - `dashboard/swiggy_sales_dashboard.xlsx`
4. Use the slicers (Month, State, Food Type, Category) to explore the data interactively
5. Refer to `Swiggy_Sales_Dashboard_Executive_Brief.pdf` for the full write-up and insights

---
<h2><a class="anchor" id="final-recommendations"></a>Final Recommendations</h2>

- Scale Saturday-specific promotions to capitalize on peak-day demand
- Strengthen delivery capacity in emerging cities (Lucknow, Hyderabad) to match Bengaluru's growth
- Expand vegetarian menu partnerships given their 64% revenue share
- Deepen partnerships with top-density merchants (KFC, McDonald's, Pizza Hut)
- Use May–August seasonal baseline for proactive demand forecasting

---
<h2><a class="anchor" id="author--contact"></a>Author & Contact</h2>

**Name- Aditya Vishwakarma**
Aspiring Business Analyst 
📧 Email: analyst.aditya09@gmail.com
📍 Address: Bhopal, Madhya Pradesh, India
🖇️ Linkedin: www.linkedin.com/in/adityaanalyst09
