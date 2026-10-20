# KB Cafe: Retail Sales Analysis & Interactive Excel Dashboard

**From 1,000 raw order lines to a one-page decision tool: finding out why a coffee business is losing orders, where its profit really comes from, and why 97% of its customers never come back.**

![Excel](https://img.shields.io/badge/Microsoft%20Excel-PivotTables%20%7C%20Slicers%20%7C%20Dashboards-217346)
![Analysis](https://img.shields.io/badge/Analysis-Retail%20%7C%20Customer%20%7C%20Profitability-6D4C41)
![Status](https://img.shields.io/badge/Status-Complete-5D7B4F)

![KB Cafe dashboard](images/01_dashboard_final.png)
*The finished one-page interactive dashboard: 4 slicers, 7 KPI tiles, 8 charts, formula-driven insights and recommended actions.*

> **Note:** GitHub cannot preview or run interactive Excel files. Download `KB_Cafe_Dashboard_v2.xlsx` to use the slicers, or read the documentation in [`docs/`](docs/README.md) for every formula and step.

---

## Table of Contents

1. [Executive Summary](#1-executive-summary)
2. [The Business and the Brief](#2-the-business-and-the-brief)
3. [Business Questions](#3-business-questions)
4. [The Data](#4-the-data)
5. [Skills and Tools Demonstrated](#5-skills-and-tools-demonstrated)
6. [Project Workflow at a Glance](#6-project-workflow-at-a-glance)
7. [Key Analytical Decisions](#7-key-analytical-decisions)
8. [Findings](#8-findings)
9. [Recommendations](#9-recommendations)
10. [Assumptions and Limitations](#10-assumptions-and-limitations)
11. [How to Use the Workbook](#11-how-to-use-the-workbook)
12. [Repository Structure](#12-repository-structure)
13. [What I Would Do Next](#13-what-i-would-do-next)
14. [About Me](#14-about-me)

---

## 1. Executive Summary

KB Cafe sells coffee in three markets (United States, Ireland, United Kingdom). I analysed **1,000 order lines (957 orders, 913 buying customers) from January 2019 to 19 August 2022** and built an interactive Excel dashboard that tells an owner, in under a minute, whether the business is growing, what earns the profit and what to do next.

### The five things the owner needs to know

| # | Finding | Evidence |
|---|---|---|
| 1 | **Sales are slipping on a like-for-like basis.** | Same-period (1 Jan to 19 Aug) sales: $7,812 (2019), $7,992 (2020), $7,620 (2021), $7,063 (2022). That is **-7.3%** vs 2021, after -4.7% the year before. |
| 2 | **The problem is traffic, not basket size.** | 2022 orders are down 9.7% (175 to 158) while average order value rose about 2.7% ($43.54 to $44.71). |
| 3 | **Sales rank and profit rank are different.** | Liberica earns **35% of profit from 27% of sales** (13.0% margin). Robusta earns **12% of profit from 20% of sales** (6.0% margin). |
| 4 | **97% of customers bought only once.** | 888 one-time vs 25 repeat customers. Repeat customers spend about **$143 vs $47**. |
| 5 | **The loyalty card shows no sign of working.** | Card holders repeat at 2.9% vs 2.6% without a card. Card holders also spend less per order ($45.08 vs $49.12). |

### The headline recommendations

1. **Win back 510 reachable lapsed customers** (642 have not ordered in over a year; 510 have an email address).
2. **Fix Robusta economics.** Lifting its margin from 6% to the 10% average adds about **$362 profit (+8%)** on the same sales.
3. **Bundle products.** Multi-item orders average **$93.69 vs $45.60**, yet only 3.2% of orders contain more than one item.
4. **Redesign the loyalty card** to reward the second purchase.
5. **Capture email at checkout** (about 20% of buyers cannot be contacted).

> Overall performance: **$45,134 revenue, $4,520 profit (10.0% margin), 957 orders, average order value $47.16, 3,627 kg sold.**

---

## 2. The Business and the Brief

KB Cafe is a small coffee retailer selling four coffee types (Arabica, Excelsa, Liberica, Robusta) in three roast levels (Light, Medium, Dark) and four pack sizes (0.2 kg, 0.5 kg, 1 kg, 2.5 kg) to customers in the United States, Ireland and the United Kingdom. Some customers hold a loyalty card.

I took the role of a **sales and business analyst** and set myself the brief an owner or investor would give:

> *"Build me something I can open in under a minute that tells me whether we are growing, where we make our money, who our customers are and what I should do about it."*

A dashboard that only describes sales would not meet that brief. The goal was to move from **describing what happened** to **supporting decisions**.

### How the project evolved

1. **Version 1** answered the basic question: how much did we sell, where and of what? It used PivotTables, four slicers, five charts and six KPI tiles.
2. **A critical review** showed three weaknesses: the trend line was not in date order, profit was invisible next to sales, and nothing said anything about customers.
3. **Version 2** fixed the chart, added 13 engineered columns, 12 PivotTables and 8 charts focused on profit, growth and customer health, and ended with written insights and actions.

---

## 3. Business Questions

| Question an owner or investor asks | Where the dashboard answers it |
|---|---|
| How big is the business and is it profitable? | KPI tiles: revenue, profit, margin, orders, AOV, kg sold |
| Are we growing against last year, fairly measured? | Same-period sales chart and scorecard |
| When do we sell most? | Monthly sales trend |
| Which products really earn the profit? | Share of sales vs share of profit; coffee economics table |
| Are we too dependent on one market? | Country mix by year |
| Which pack sizes carry the revenue? | Sales by pack size |
| Do customers come back? | Repeat vs one-time customers; repeat-rate tile |
| Is the loyalty card working? | Repeat rate: card holders vs no card |
| How many customers have we lost, and can we reach them? | Customer recency chart (with email coverage) |
| What should we do next? | Insights and recommended actions boxes |

---

## 4. The Data

Three related tables: `orders` (1,000 order lines), `customers` (1,000 registered customers) and `products` (48 products). Orders link to customers and products through `Customer ID` and `Product ID`.

| Item | Detail |
|---|---|
| **Period** | 2 January 2019 to 19 August 2022 |
| **Markets** | United States, Ireland, United Kingdom |
| **Products** | 4 coffee types, 3 roast levels, 4 pack sizes (48 products) |
| **Grain** | One row per product on an order (an *order line*) |

**Data realities that shaped the analysis**

- One row is not one order: 1,000 rows but **957 orders** (74 rows belong to 31 multi-item orders).
- One row is not one customer: **913 buying customers**. 87 registered customers never bought (70 of them have an email address).
- Not every customer has an email: 726 of 913 buyers (79.5%) can be reached.
- Profit is only available as a **per-unit figure** in the products table, so total profit is quantity times unit profit.

The full data model, the lookups I built and the data-quality notes are in [docs/01_data_and_preparation.md](docs/01_data_and_preparation.md).

---

## 5. Skills and Tools Demonstrated

| Skill | How it shows up in this project |
|---|---|
| **Data preparation** | Joined three tables with `INDEX`/`MATCH`, handled blank emails, converted product codes to readable names via a lookup list |
| **Feature engineering** | Built 13 helper columns to turn order lines into orders, customers, recency, baskets and like-for-like periods |
| **Structured table formulas** | Excel Tables with structured references, expanding-range `COUNTIF`, `SUMIFS`, `MAXIFS`, named formulas |
| **Pivot analysis** | 12 PivotTables with calculated fields, Show Values As (percent of total, difference from previous) and report filters |
| **Metric design** | Defined revenue, profit, margin, AOV, profit per kg, repeat rate, recency segments and same-period growth |
| **Interactive dashboard design** | Slicers with a deliberate connection matrix, GETPIVOTDATA-driven KPI tiles, PivotCharts and formula-driven charts |
| **Visual design** | Consistent palette, 21-column grid layout, one-page landscape print setup |
| **Insight writing** | Formula-driven insight sentences that update with filters; "what, so what, now what" structure |
| **Validation** | Checked every helper column and every slicer state against known totals |
| **Business communication** | Translated findings into sized, measurable recommendations and stakeholder-ready language |

**Tools:** Microsoft Excel (PivotTables, PivotCharts, slicers, Tables, Data Validation, Name Manager, conditional formatting).

---

## 6. Project Workflow at a Glance

| Step | What I did | Detail |
|---|---|---|
| 1. Understand | Wrote the business questions before touching a formula | [Section 3](#3-business-questions) |
| 2. Prepare | Filled the orders table with lookups to customers and products; made it an Excel Table | [docs/01](docs/01_data_and_preparation.md) |
| 3. Engineer | Added 13 helper columns and one named formula (`DataEnd`) | [docs/02](docs/02_feature_engineering.md) |
| 4. Analyse | Built 12 PivotTables, one per question or chart | [docs/03](docs/03_pivot_analysis.md) |
| 5. Visualise | Built 8 charts (7 PivotCharts, 1 formula-driven) | [docs/04](docs/04_dashboard_design.md) |
| 6. Interact | Added 4 slicers, each connected deliberately | [docs/04](docs/04_dashboard_design.md) |
| 7. Design | Laid out a one-page, 21-column dashboard | [docs/04](docs/04_dashboard_design.md) |
| 8. Narrate | Wrote formula-driven insights and actions | [docs/04](docs/04_dashboard_design.md) |
| 9. Validate | Checked helper columns, KPIs and slicer states against known totals | [docs/05](docs/05_validation_and_decisions.md) |

A click-by-click build guide is in [docs/KB_Cafe_Dashboard_v2_Build_Guide.docx](docs/KB_Cafe_Dashboard_v2_Build_Guide.docx).

---

## 7. Key Analytical Decisions

The choices that most changed the results (full table in [docs/05](docs/05_validation_and_decisions.md)):

- **Count distinct orders with a flag, not rows.** Rows say 1,000 orders; the truth is 957.
- **Classify repeat customers by distinct orders, not rows.** Counting rows would label **52** customers as repeat; the correct number is **25**.
- **Compare the same period each year** (1 Jan to 19 Aug), because 2022 is incomplete.
- **Measure recency to the last date in the data, not `TODAY()`**, otherwise every customer would look lapsed.
- **Keep the Year slicer off the customer charts**, because the customer flag sits on each customer's first order.

---

## 8. Findings

### 8.1 Growth: a decline driven by fewer orders

| Year | Same-period sales (1 Jan to 19 Aug) | Orders | Avg order value | Growth |
|---|---|---|---|---|
| 2019 | $7,812 | 155 | $50.40 | |
| 2020 | $7,992 | 168 | $47.57 | +2.3% |
| 2021 | $7,620 | 175 | $43.54 | -4.7% |
| 2022 | $7,063 | 158 | $44.71 | **-7.3%** |

Revenue is orders times average order value. In 2022 orders fell 9.7% while AOV rose about 2.7%, so the decline is a **traffic problem, not a basket-size problem**. AOV is still 11% below 2019.

*Seasonality caveat:* September to December is the peak (34%, 34% and 42% of sales in 2019 to 2021), and 2022's peak is still to come. This is an **early warning, not a verdict**.

### 8.2 Profitability: sales rank is not profit rank

| Coffee | Share of sales | Share of profit | Margin | Profit per kg |
|---|---|---|---|---|
| Liberica | 27% | **35%** | 13.0% | $1.89 |
| Excelsa | 27% | 30% | 11.0% | $1.51 |
| Arabica | 26% | 23% | 9.0% | $1.02 |
| Robusta | 20% | **12%** | **6.0%** | $0.63 |

Each dollar of Robusta earns about half what a dollar of Liberica earns. Roast level makes almost no difference (margins of about 10% across Light, Medium and Dark), which is why I removed the Coffee x Roast chart from the dashboard.

### 8.3 Customers: the business has a retention problem

- **97.3% of customers (888 of 913) bought once.**
- The 25 repeat customers are 2.7% of customers but **7.9% of revenue**, and spend about **$143 versus $47** per customer.
- **Loyalty card:** card holders repeat at 2.9% (13 of 443) versus 2.6% (12 of 470). There is **no evidence the card changes behaviour**. Card holders also spend less per order ($45.08 versus $49.12).

*Caution:* counts of 13 and 12 are small, so I phrase it as "no evidence it works", not "it does not work".

### 8.4 Customer base health

| Segment | Customers | With email |
|---|---|---|
| Active (0 to 90 days) | 57 | 41 |
| At risk (91 to 365 days) | 214 | 175 |
| **Lapsed (over 365 days)** | **642** | **510** |
| Total | 913 | 726 |

Seventy percent of customers have not returned in over a year, and 79.5% of buyers can be contacted by email at almost no cost. A further 87 registered customers never bought at all (70 have an email address).

### 8.5 Baskets and pack sizes

- Only **31 of 957 orders (3.2%)** contain more than one product, but they average **$93.69 versus $45.60**.
- The 2.5 kg bag is **53%** of sales. The 0.2 kg bag is about 25% of units but only about 7% of revenue, so it may cost more to handle than it earns (cost data would confirm).

### 8.6 Markets

- The **United States is 79%** of sales overall and **81% in 2022**.
- The **United Kingdom's share fell from 8% (2019) to 2% (2022)**; same-period UK sales dropped from $289 to $130 between 2021 and 2022.
- Ireland grew from 15% to 17% of sales.

### 8.7 What did not matter (and why I left it out)

- **Top 5 customers:** together only 3.3% of sales, so no single customer matters. I removed the chart.
- **Coffee x Roast:** all roasts earn about the same margin, so it rarely changes a decision.

---

## 9. Recommendations

| # | Action | Why | Estimated impact | Success measure |
|---|---|---|---|---|
| 1 | **Win-back email campaign** to lapsed customers | 510 reachable lapsed customers; repeat buyers are worth about 3 times more | *Illustration:* a 5% response (about 25 people) at $47 each is about $1,200 revenue, roughly $120 profit at a 10% margin | Response rate, orders from the list |
| 2 | **First-purchase offer** to 87 registered non-buyers | They signed up but never converted; 70 have email | Small, near-zero cost test | Sign-up to first-purchase conversion |
| 3 | **Review Robusta cost or price** | 6% margin versus 10% average | About **$362 extra profit (+8%)** on unchanged volume | Margin by coffee after the change |
| 4 | **Bundle products** | Multi-item orders average about double | Lifts AOV if the share of multi-item orders rises | Multi-item share, AOV |
| 5 | **Redesign the loyalty card** to reward a second purchase | No repeat lift from the current card | Moves repeat rate toward a target (for example 5%) | Repeat rate, card holders vs others |
| 6 | **Capture email at checkout** | About 20% of buyers have no email (187 of 913) | Expands the reachable audience for actions 1 and 5 | Email coverage % |

**Be realistic about scale.** The business sold about $45k over roughly 3.6 years, so these are modest, real improvements, and the first two cost almost nothing. The impact figures are **illustrations with stated assumptions**, not forecasts. For example, the Robusta estimate assumes volume does not fall if price rises, which is why the action says "review cost or price".

---

## 10. Assumptions and Limitations

- **Currency** is shown as `$` because the dataset does not state one.
- **Profit** uses the per-unit profit figure supplied in the products table. Marketing, handling and overhead costs are not in the data.
- **Recency is measured as of 19 August 2022**, the last date in the data, not today's date.
- **Recency thresholds (90 and 365 days)** are business judgements and can be adjusted.
- **The 2022 comparison is year-to-date.** The peak season (September to December) is not yet in the data, so the decline is an early warning, not a final result.
- **Small counts:** only 25 customers are repeat buyers, so retention percentages are directional.
- **Findings describe what happened;** they suggest causes but do not prove them. Recommendations are hypotheses to test, and impact figures are illustrations.
- **Customer charts are all-time snapshots** and do not respond to the Year slicer by design.

---

## 11. How to Use the Workbook

1. Open `KB_Cafe_Dashboard_v2.xlsx` in **Microsoft Excel 2019 or Microsoft 365**. Click **Enable Editing** if it opens in Protected View.
2. Go to the **Dashboard** sheet.
3. Click the **slicers** (Year, Country, Loyalty Card, Coffee Type) to filter. Use **Ctrl+Click** to select several items, or the funnel icon to clear.
4. Read the dashboard top to bottom: KPIs, trend, products, customers, insights, actions.
5. After changing the data, use **Data > Refresh All**.

| Sheet | Contents |
|---|---|
| `Dashboard` | The one-page interactive dashboard |
| `Analysis` | PT-01 to PT-12 |
| `Insights` | Insight and action formulas |
| `Coffee_data` (orders) | Cleaned orders table with 13 helper columns |
| `customers`, `products` | Source reference tables |

---

## 12. Repository Structure

```
kb-cafe-sales-dashboard/
|-- README.md
|-- KB_Cafe_Dashboard_v2.xlsx
|-- data/
|   |-- README.md                              (source, licence and privacy note)
|   `-- KB_Cafe_Retail_Sales_Data_Raw.xlsx     (add only if you have the right to share it)
|-- images/
|   |-- 01_dashboard_final.png                 (screenshot of the finished dashboard)
|   |-- 02_dashboard_blueprint.png
|   |-- 03_trend_sort_fix.png
|   |-- 04_target_design_mockup.png
|   `-- pivots/
|       `-- pt01.png ... pt12.png
`-- docs/
    |-- README.md
    |-- 01_data_and_preparation.md
    |-- 02_feature_engineering.md
    |-- 03_pivot_analysis.md
    |-- 04_dashboard_design.md
    |-- 05_validation_and_decisions.md
    `-- KB_Cafe_Dashboard_v2_Build_Guide.docx
```

---

## 13. What I Would Do Next

- **Add targets** (for example repeat rate 10%, margin 12%) and show them as traffic-light indicators beside each KPI tile. A KPI without a target is only a number.
- **Cohort analysis:** group customers by first-purchase month and track how many return, to see whether retention is improving or worsening.
- **Forecast the peak season** to estimate where 2022 may finish.
- **Customer lifetime value** and acquisition-cost analysis once marketing spend data is available.
- **A/B test the win-back offer** on a random half of lapsed customers to measure true uplift.
- **Rebuild the same analysis in Power BI and SQL** to compare approaches (DAX measures and SQL window functions for the repeat and recency logic).

---

## 14. About Me

**[Your Name]**: Data Analyst

I turn raw business data into decisions: cleaning and modelling data, building dashboards people actually use, and explaining the findings in plain language.

- LinkedIn: [your-link]
- GitHub: [your-link]
- Portfolio: [your-link]
- Email: [your-email]
