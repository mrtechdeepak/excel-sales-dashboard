# Excel Sales Dashboard - Demo Instructions

## Quick Start Demo

Follow these steps to see the Excel Sales Dashboard in action:

### Step 1: Install Dependencies
```bash
pip install -r requirements.txt
```

### Step 2: Generate the Dashboard
```bash
python generate_sales_dashboard.py
```

This creates `Sales_Dashboard.xlsx` with sample data.

### Step 3: Open the Dashboard
Open `Sales_Dashboard.xlsx` in Microsoft Excel or Google Sheets.

---

## Dashboard Tour

### Sheet 1: Raw Data
- Contains 240 sales records across 12 months
- Columns: Order ID, Date, Month, Region, Product Category, Sales Channel, Units Sold, Unit Price, Sales, Cost, Profit, Target, Discount %
- Fully formatted with headers and borders

### Sheet 2: Dashboard (Interactive)

#### Filter Controls (Top)
- **Selected Region**: Dropdown filter (All, North, South, East, West, Central)
- **Selected Month**: Dropdown filter (All, Jan-Dec)
- Change these dropdowns to see all charts and KPIs update dynamically

#### KPI Cards (Middle Left)
- **Total Sales**: Sum of all sales for selected filters
- **Total Profit**: Sum of profit margin
- **Units Sold**: Total units moved
- **Target Achieved**: Percentage of sales target met

#### Summary Tables
1. **Monthly Summary** (Left): Sales and Profit by month
2. **Regional Summary** (Center): Sales and Profit by region
3. **Category Summary** (Right): Sales and Profit by product category

#### Interactive Charts
1. **Monthly Sales Trend** (Line Chart): Shows sales progression across months
2. **Regional Sales** (Bar Chart): Compares sales performance by region
3. **Sales by Category** (Pie Chart): Visualizes sales distribution

---

## Try These Demo Actions

1. **Filter by Region**: Click cell B4 and select "North" → All charts update
2. **Filter by Month**: Click cell B5 and select "May" → KPIs recalculate
3. **Dual Filter**: Select "East" region AND "Jun" month → Dashboard shows filtered data
4. **Reset**: Click B4 and B5, set both to "All" → See full year overview

---

## Key Features

✅ **Fully Dynamic** - All formulas use SUMIFS with IF conditions for real-time filtering
✅ **Professional Styling** - Color-coded headers, borders, and KPI cards
✅ **Interactive Charts** - Line, bar, and pie charts respond to filter changes
✅ **Scalable Data** - Easy to add more sales records; dashboard auto-updates
✅ **Conditional Formatting** - Green highlighting for positive metrics
✅ **Frozen Headers** - Easy navigation even with large datasets

---

## Data Included

- **Regions**: North, South, East, West, Central
- **Product Categories**: Electronics, Furniture, Office Supplies, Accessories
- **Sales Channels**: Online, Retail, Distributor, Wholesale
- **Time Period**: 12 months of data (January - December 2024)
- **Total Records**: 240 sales transactions

---

## Customize the Data

Edit `generate_sales_dashboard.py` to change:
- Product categories, regions, or sales channels
- Sales values or profit margins
- Chart types or styling
- KPI formulas or summary metrics

Then regenerate with:
```bash
python generate_sales_dashboard.py
```

---

## Requirements

- Python 3.7+
- openpyxl 3.1.5
- Excel or compatible spreadsheet application

---

Enjoy exploring the dynamic dashboard! 📊
