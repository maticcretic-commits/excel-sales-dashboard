# Excel Sales Dashboard 📊

[![GitHub stars](https://img.shields.io/github/stars/maticcretic-commits/excel-sales-dashboard?style=social)](https://github.com/maticcretic-commits/excel-sales-dashboard/stargazers)
[![Last commit](https://img.shields.io/github/last-commit/maticcretic-commits/excel-sales-dashboard)](https://github.com/maticcretic-commits/excel-sales-dashboard/commits/main)
[![Cost: Free](https://img.shields.io/badge/cost-%E2%82%B90-brightgreen)](https://github.com/maticcretic-commits/excel-sales-dashboard)

> **A realistic 500-row sales dataset turned into a polished, formula-driven Excel dashboard. Zero hardcoded numbers — every KPI, table, and chart reads live from the data.**

Built for freelancing clients who live in spreadsheets: open the file, change any
row, and the whole dashboard recalculates.

## What's inside

| File | What it is |
|---|---|
| `sales-dashboard.xlsx` | The dashboard — open in Excel or LibreOffice |
| `generate_dashboard.py` | Python script that builds the workbook (so the dataset is reproducible) |

### The workbook

- **`Data` sheet** — 500 sales records (Jan–Dec 2025): OrderID, Date, Region,
  SalesRep, Product, Category, Units, UnitPrice, Revenue, Cost, Profit.
- **`Dashboard` sheet** — KPI cards, monthly trend table, region and product
  breakdowns, and two native Excel charts.

## Excel skills demonstrated

- **Formulas (all live, none hardcoded):** `SUMIFS`, `COUNTIFS`, `AVERAGEIFS`-style
  logic, `XLOOKUP`, `EDATE`, `IF` guards against divide-by-zero
- **Native charts:** monthly revenue line chart, revenue-by-region column chart —
  built with Excel chart objects, not pasted images
- **Data validation:** Region and Category dropdown lists on the Data sheet
- **Conditional formatting:** big deals (₹1L+) highlighted green, thin-margin rows
  flagged red
- **Presentation:** frozen panes, auto-filter, print-ready fit-to-page, ₹ number
  formats, KPI card layout

## How to use

1. Download `sales-dashboard.xlsx` and open it in Excel (desktop or web) or LibreOffice.
2. Edit any row on the **Data** sheet — the Dashboard recalculates instantly.
3. To regenerate the dataset with fresh random data: `python3 generate_dashboard.py`
   (needs `openpyxl`).

## For clients

This is the template I use for sales-performance, expense-tracking, and
inventory dashboards: your raw data in, a boardroom-ready dashboard out —
formulas you can audit, charts that update themselves, no macros required.

## License

MIT — use it, fork it, sell services built on it.

## ☕ Support my work
If this project was useful, you can support it with Bitcoin: `bc1q6q75k8zjxvw7w02lmdprpy6xx6qk4lzz2rmvay`
