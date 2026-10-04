# Excel + Power BI Sales Reporting Suite

A professional-grade sales analytics solution combining Excel's powerful formulas with Power BI export capabilities for enterprise-level reporting.

## 📊 Architecture Overview

This suite delivers:
- **Excel Dashboard**: Interactive KPI tracking, regional analysis, and trend forecasting
- **Power BI Ready**: Structured data export with dimensional modeling
- **Real-time Sync**: Automated refresh mechanisms for live performance monitoring
- **Executive Reports**: Pre-built summary views for C-suite analytics

---

## 🚀 Quick Start

### 1. Generate Sales Data & Excel Dashboard
```bash
pip install -r requirements.txt
python generate_sales_dashboard.py
```

### 2. Output Files
- `Sales_Dashboard.xlsx` - Interactive Excel workbook
- `sales_data_export.csv` - Power BI data source (auto-generated)

### 3. Connect to Power BI
- Open Power BI Desktop
- Select **Get Data → Excel**
- Import `Sales_Dashboard.xlsx` → Raw Data sheet
- Build visualizations from the structured dataset

---

## 📈 Excel Dashboard Components

### Performance Dashboard Sheet
Located in the "Dashboard" tab with five integrated sections:

#### 1. Filter Control Panel (Top)
```
Selected Region: [Dropdown: All/North/South/East/West/Central]
Selected Month:  [Dropdown: All/Jan-Dec]
Selected Category: [Dropdown: All/Electronics/Furniture/Office Supplies/Accessories]
```
- Real-time cascade filtering
- Multi-select capability for Power BI export

#### 2. Executive KPI Section
| Metric | Formula | Business Use |
|--------|---------|--------------|
| **Total Sales** | SUMIFS across filtered region/month | Revenue tracking |
| **Gross Profit** | Cost variance analysis | Margin health |
| **YoY Growth %** | Month-over-month comparison | Trend analysis |
| **Target % Achievement** | Actual vs. planned sales | Forecast accuracy |

#### 3. Time Series Analysis
- **Monthly Sales Trend**: Line chart with trend line
- **Profit Margin Trend**: Dual-axis showing margin trajectory
- **Forecast Zone**: 30/60/90-day projections

#### 4. Dimensional Analysis
- **Regional Heat Map**: Sales performance by geography
- **Category Breakdown**: Product line contribution analysis
- **Channel Performance**: Online vs. Retail vs. Distributor vs. Wholesale

#### 5. Data Quality Indicators
- Missing data alerts
- Outlier detection (red flags for anomalies)
- Data freshness timestamp

---

## 🔗 Power BI Integration Strategy

### Data Flow Architecture
```
Excel Raw Data Sheet
         ↓
    CSV Export
         ↓
  Power BI Desktop
         ↓
  Data Modeling (Fact/Dim Tables)
         ↓
  DAX Calculations
         ↓
  Interactive Visualizations
         ↓
  Power BI Service (Cloud Publishing)
```

### Dimensional Model (Power BI)

**Fact Table: Sales**
- OrderID, DateKey, RegionKey, CategoryKey, ChannelKey
- Sales, Cost, Profit, Units, Discount
- Target, ActualTarget

**Dimension Tables:**
- **Date Dim**: Month, Quarter, Year, Fiscal Period, Holiday Flag
- **Region Dim**: Region, SubRegion, Territory, Manager
- **Category Dim**: Category, SubCategory, ProductLine
- **Channel Dim**: Channel, ChannelType, Partner

### Sample DAX Calculations
```dax
Total Sales = SUM(Sales[Sales])

Sales YTD = CALCULATE([Total Sales], 
    DATESYTD(DateDim[Date]))

Sales Growth % = DIVIDE(
    [Total Sales] - CALCULATE([Total Sales], 
        DATEADD(DateDim[Date], -12, MONTH)),
    CALCULATE([Total Sales], 
        DATEADD(DateDim[Date], -12, MONTH)))

Target Achievement % = DIVIDE([Total Sales], [Total Target])

Profit Margin % = DIVIDE([Total Profit], [Total Sales])
```

---

## 📊 Recommended Power BI Visualizations

### Page 1: Executive Summary
- **Card**: Total Sales (Current Period)
- **Card**: Profit Margin %
- **Card**: Target Achievement %
- **KPI Visual**: Sales vs. Target
- **Line Chart**: Sales Trend (Last 12 Months)

### Page 2: Regional Performance
- **Map Visual**: Sales by Region (bubble size = sales volume)
- **Bar Chart**: Regional Ranking
- **Table**: Region Details (Sales, Profit, Margin, Growth %)
- **Slicers**: Region, Month, Category

### Page 3: Product Analysis
- **Pie Chart**: Sales by Category
- **Clustered Bar**: Category Profitability
- **Scatter**: Price vs. Volume Analysis
- **Table**: Top 20 Products

### Page 4: Channel Analytics
- **Stacked Bar**: Sales by Channel over Time
- **Funnel**: Order → Delivery → Return Rate
- **Table**: Channel Metrics (AOV, Conversion, Margin)

### Page 5: Forecasting & Trends
- **Decomposition Tree**: Sales drivers (Region → Category → Channel)
- **Line with Forecast**: 12-month projection with confidence interval
- **Ribbon Chart**: Top performers over time
- **Q&A Visual**: Natural language queries

---

## 🔄 Workflow: Excel ↔ Power BI

### Option 1: Manual Refresh
1. Run `python generate_sales_dashboard.py` to update Excel
2. In Power BI → Refresh data
3. Dashboards auto-update

### Option 2: Automated Sync (Advanced)
```python
# Add to generate_sales_dashboard.py
import subprocess

# After Excel generation, export to CSV for Power BI
def export_to_csv():
    df = pd.read_excel('Sales_Dashboard.xlsx', sheet_name='Raw Data')
    df.to_csv('sales_data_export.csv', index=False)
    print("Data exported to sales_data_export.csv for Power BI")
```

### Option 3: Cloud Integration
- Upload `Sales_Dashboard.xlsx` to SharePoint
- Connect Power BI to SharePoint folder
- Enable scheduled refresh (daily/hourly)

---

## 📋 Data Dictionary

| Column | Type | Description | Power BI Use |
|--------|------|-------------|--------------|
| Order ID | Integer | Unique transaction identifier | Grain level |
| Date | Date | Transaction date | Time dimension key |
| Month | Text | Calendar month | Slicer/grouping |
| Region | Text | Sales region | Geographic dimension |
| Product Category | Text | Product line | Dimensional filter |
| Sales Channel | Text | Sales method (Online/Retail/etc) | Channel analysis |
| Units Sold | Integer | Quantity sold | Volume metric |
| Unit Price | Decimal | Price per unit | Revenue composition |
| Sales | Decimal | Total revenue | Primary KPI |
| Cost | Decimal | COGS | Margin calculation |
| Profit | Decimal | Sales - Cost | Profitability KPI |
| Target | Decimal | Sales target | Achievement %, forecasting |
| Discount % | Decimal | Discount applied | Promotion impact |

---

## 🎯 Use Cases

### Sales Manager Scenario
1. Open Excel Dashboard
2. Filter by Region = "North"
3. Identify underperforming months
4. Switch to Power BI for deep-dive
5. Create custom report for weekly standup

### CFO Scenario
1. Pull Power BI executive report
2. Compare actual vs. target
3. Analyze margin trends
4. Export to board presentation
5. Schedule automated refresh

### Marketing Team Scenario
1. Use Excel to track channel performance
2. Export to Power BI for campaign ROI analysis
3. Build attribution model
4. Present customer acquisition insights

---

## 🛠 Customization Guide

### Add New Metrics to Excel
Edit `generate_sales_dashboard.py`:
```python
# Add to KPI section
"Return Rate": "=COUNTIFS(RawData, criteria)/COUNTA(RawData)",
"Customer Lifetime Value": "=AVERAGE sales per customer",
"Days Sales Outstanding": "=AVERAGE Accounts Receivable / Daily Sales"
```

### Extend Power BI Model
1. Add new columns to Raw Data sheet
2. Refresh Power BI connection
3. Create new measures in Data Model
4. Build new visualizations

### Schedule Automated Reports
Power BI → Settings → Scheduled Refresh → Daily at 6 AM

---

## 📦 Required Tools

| Tool | Version | Purpose |
|------|---------|---------|
| Python | 3.7+ | Dashboard generation |
| openpyxl | 3.1.5+ | Excel creation |
| Microsoft Excel | 2019+ | Interactive analysis |
| Power BI Desktop | Latest | Advanced visualization |
| Power BI Service | Cloud | Publishing & sharing |

---

## 🚀 Deployment Checklist

- [ ] Generate Excel dashboard
- [ ] Export data to CSV
- [ ] Create Power BI data model
- [ ] Build executive dashboard
- [ ] Set up scheduled refresh
- [ ] Share with stakeholders
- [ ] Document filters & calculations
- [ ] Set up automated alerts

---

## 📞 Support & Enhancement

**For Excel Issues:**
- Check filter cell references (B4, B5, B6)
- Verify SUMIFS formulas use correct sheet names
- Ensure Raw Data table is not corrupted

**For Power BI Issues:**
- Refresh data source connection
- Clear Power BI cache
- Verify DAX calculations for circular references
- Check for mismatched data types

**To Add Features:**
- Predictive analytics (Python ML models)
- Sentiment analysis on customer feedback
- Anomaly detection for fraud prevention
- Real-time data ingestion from APIs

---

## 📄 Sample Report Output

**Executive Summary Report (Monthly)**
```
Sales Performance Report - October 2024

KPIs:
  • Total Sales: $2.8M (+15% vs prior month)
  • Gross Profit: $840K (30% margin)
  • Target Achievement: 105% ✓
  • Top Region: East ($650K, +22% growth)
  
Risks:
  ⚠ Office Supplies declining (-8% MoM)
  ⚠ Distributor channel underperforming
  
Opportunities:
  ✓ Online sales accelerating (+35%)
  ✓ Electronics strong growth trajectory
```

---

Enjoy professional-grade sales reporting! 📊✨
