# Regional Sales Performance Dashboard

A simple, business-friendly Power BI dashboard that helps a **Sales Manager** understand how the company is performing — without needing any technical or data skills.

This project was built for the **Data Analyst Assignment** (KS AI & Cloud Solutions). It turns messy spreadsheet data into one clear story: *Is the business growing? Are managers hitting their targets? Are returns and delivery delays hurting customers?*

---

## Who is this for?

This dashboard is designed for:

- **Sales Managers** who run monthly business review meetings  
- **Regional Managers** who need to see how their region and product categories are doing  
- **Leadership** who want quick answers, not complicated reports  

---

## What problem does this solve?

Before this dashboard, the company had:

- Sales data spread across multiple Excel/CSV files  
- No easy way to compare **this year vs last year**  
- No clear view of whether each **Regional Manager** is meeting **category targets**  
- No single place to see **returns** and **delivery delays**  

This dashboard answers questions like:

| Question | Where to look |
|----------|----------------|
| Is sales growing compared to last year? | Page 1 — Sales YoY % |
| Is profit growing? | Page 1 — Profit YoY % |
| Which region sells the most? | Page 1 — Sales by Region |
| Is each manager meeting their category targets? | Page 2 — Manager & Category table |
| Are returns going up or down? | Page 3 — Returns YoY % |
| Are deliveries often late? | Page 3 — Delay Rate % |
| Where should management take action? | Page 3 — Action Areas table |

---

## Project files (what’s in this folder)

| File / Folder | What it is |
|---------------|------------|
| **Sales_Performance_Dashboard.pbix** | The main Power BI dashboard — **open this file** |
| **order_2015.csv**, **order_2016.csv**, **order_2017.csv** | Raw sales order data (3 years) |
| **People.csv** | Which Regional Manager is responsible for each region |
| **Returns.csv** | Which orders were returned by customers |
| **Target.csv** | Company sales targets by product category and year |
| **processed/** | Cleaned data files used inside the dashboard |
| **build_dashboard.py** | Script that cleans data and builds the dashboard (for technical users) |

---

## The dashboard story (3 pages)

The dashboard is organized like a **business story** — from big picture to details to risks.

### Page 1 — Business Health Overview

**Purpose:** *“How is the business doing overall?”*

**What you see:**

- **KPI cards** at the top: Total Sales, Sales Growth (YoY), Total Profit, Profit Growth, Orders, Target Achievement, Profit Margin  
- **Sales by Region** — which region contributes the most  
- **Target by Region** — what each region was expected to sell  
- **Monthly Sales Trend** — how sales moved month by month  
- **Sales by Category** — Furniture, Office Supplies, Technology mix  
- **Target Achievement by Category** — which product lines are above or below target  
- **Sales by Regional Manager** — quick manager comparison  

**Filters:** Year, Region, Category

---

### Page 2 — Managers and Targets

**Purpose:** *“Are Regional Managers hitting their category targets?”*

This page answers the assignment’s key requirement: a manager can do well in one category but fail in another. The table shows **each manager × each category** so gaps are easy to spot.

**What you see:**

- **Sales Target**, **Total Sales**, **Target Gap**, **Target Achievement %**  
- **Detail table** with: Region, Regional Manager, Category, Sales, Target, Achievement %, Gap, and **Target Status** (Target Met / Below Target)  
- **Charts:** Achievement by Manager, Target Gap by Category  

**Filters:** Year, Regional Manager  

**How to read Target Status:**

- **Target Met** = Sales reached or beat the target for that manager + category  
- **Below Target** = Sales fell short — this is where coaching or action may be needed  

---

### Page 3 — Returns, Delays and Actions

**Purpose:** *“Are returns and late deliveries hurting the business? Where should we focus?”*

**What you see:**

- **Return metrics:** Returned Orders, Returns YoY %, Return Rate %  
- **Delivery metrics:** Delayed Orders, Delays YoY %, Delay Rate %, Average Ship Delay Days  
- **Charts:** Return rate by Region and Category; Delay rate by Ship Mode; Average delay by Manager  
- **Action Areas table** — combines Target Achievement, Return Rate, Delay Rate, and Target Status so leadership can see **high-risk combinations** in one place  

**Filters:** Year, Region  

---

## Key terms explained (plain English)

| Term | Meaning |
|------|---------|
| **YoY (Year over Year)** | Compare this year to the previous year. Example: 2017 vs 2016. A positive % means growth; negative means decline. |
| **Sales Target** | How much the company expected to sell in a region + category for a given year. |
| **Target Achievement %** | Actual sales ÷ target. **100%** = exactly on target. **Above 100%** = beat target. **Below 100%** = missed target. |
| **Target Gap** | Target minus actual sales. **Negative gap** often means you sold *more* than target (good). **Positive gap** can mean you’re still short of target. |
| **Return Rate %** | Share of orders that were returned. Higher rate may mean product or service issues. |
| **Delay Rate %** | Share of orders shipped later than expected (based on ship mode rules). |
| **Regional Manager** | Person responsible for sales in their region (from People.csv). |
| **Category** | Product type: Furniture, Office Supplies, or Technology. |

---

## What we did to the data (data preparation)

All raw files were cleaned and combined before building the dashboard.

### Step 1 — Combined orders

- Merged **order_2015**, **order_2016**, and **order_2017** into one **Orders** table (6,472 order lines).

### Step 2 — Fixed dates

- Order dates in the CSV were stored as Excel numbers; we converted them to real dates (e.g. 2015-06-02).

### Step 3 — Added Regional Manager

- Joined **People.csv** so each order knows its manager (North → Ross DeVincentis, Central → Emily Burns, South → Damala Kotsonis).

### Step 4 — Marked returns

- Joined **Returns.csv** (removed duplicates) and flagged each order as **Returned: Yes / No**.

### Step 5 — Calculated delivery delays

- **Ship Delay Days** = days between order date and ship date  
- **Delayed?** = Yes if actual delay was longer than expected for that ship mode:

| Ship Mode | Expected max delay |
|-----------|-------------------|
| Same Day | 0 days |
| First Class | 1 day |
| Second Class | 2 days |
| Standard Class | 4 days |

### Step 6 — Built targets

- **Target.csv** gives **company-wide** targets per **category** and **year** (e.g. Furniture 2017 = 260K).  
- **Assumption:** Targets are split across the 3 regions (**North, Central, South**) based on each region’s **share of actual sales** in that category and year.  
  - Example: If Central sells 60% of Furniture in 2017, Central gets ~60% of the Furniture target.  
  - This gives **different targets per region/manager**, which matches real business accountability better than splitting equally.

### Step 7 — Created measures (calculated KPIs)

Examples: Total Sales, YoY %, Target Achievement %, Return Rate %, Delay Rate %, Target Status.

---

## Important assumptions (say this in your presentation)

1. **Targets are regional**, derived from company category targets in Target.csv.  
2. **Default year** for some calculations is **2017** when no year is selected — always pick a **Year** in the slicer for accurate numbers.  
3. **YoY** compares the selected year to the year before (e.g. 2017 vs 2016).  
4. **Returns** are matched at order level; duplicate return rows in the source were removed.  
5. The dashboard is meant for **monthly business reviews**, not daily operations.

---

## How to open and use the dashboard

### Requirements

- **Microsoft Power BI Desktop** (free download from Microsoft)  
- Windows PC (Power BI Desktop is mainly for Windows)

### Steps

1. Download or clone this repository.  
2. Open **`Sales_Performance_Dashboard.pbix`** in Power BI Desktop.  
3. On each page, use the **slicers** (filters on the left/top):  
   - Select **Year** (recommended: **2017** for latest full-year view)  
   - Optionally select **Region**, **Category**, or **Regional Manager**  
4. Read KPI cards first, then charts, then detail tables.  
5. For presentations: start on **Page 1**, then **Page 2** for manager gaps, then **Page 3** for risks.

### Tips for non-technical users

- Click a bar in a chart to filter other visuals on the same page.  
- Click again to clear the filter.  
- If numbers look odd, check that **Year** is selected on the slicer.  
- **Green / good** and **red / bad** in YoY cards depend on context — growth in sales is usually good; growth in returns or delays is usually bad.


## For technical users (optional)

### Rebuild the dashboard from source

```bash
pip install pandas pbix-mcp
python build_dashboard.py
```

This will:

1. Run ETL and save cleaned CSVs to `processed/`  
2. Generate `Sales_Performance_Dashboard.pbix`  
3. Apply light theme formatting and measure formats  

### Data model (simple view)

```

Orders (fact table)
    └── linked to Targets via Target Key (Region + Category + Year)

KPI Metrics (hidden table holding all DAX measures)
```
<img width="1920" height="1080" alt="image" src="https://github.com/user-attachments/assets/2a324dcc-c867-4e92-98f8-e3df9d4d2061" />
