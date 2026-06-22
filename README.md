# Norman Inventory Dashboard

An Excel-based inventory tracking system for window treatment products, built to parse, validate, and report on product data using formulas only — no hardcoded values.

## Overview

This workbook simulates a product inventory dataset for a window-covering manufacturer, tracking units produced, warranty coverage, and product lifecycle data across multiple brands and product lines.

## What it does

- **Parses structured Product IDs** (e.g. `LV96RMN020`) into Brand, Product Line, Manufacture Year, and a generated Replacement Product ID using `LEFT`, `MID`, `RIGHT`, and `CONCATENATE`
- **Looks up full brand and product line names** from reference tables using `VLOOKUP`
- **Calculates product age** dynamically based on a two-digit manufacture year, handling the year-rollover edge case with nested `IF` logic
- **Flags warranty coverage status** (Covered / Not Covered) by comparing units produced against warranty thresholds
- **Summarizes data** by product line, fabric color, and coverage status
- **Visualizes the summary** with a bar chart, pie chart, and donut chart built directly from the underlying formulas

## Sheets

| Sheet | Purpose |
|---|---|
| `Sheet1` | Dashboard with summary tables and charts |
| `norman_inventory` | Raw product data with all formulas (52 products, 14 columns) |

## Skills demonstrated

- Excel formulas: `VLOOKUP`, `LEFT`, `MID`, `RIGHT`, `IF`, `CONCATENATE`, `UPPER`
- Data validation and conditional logic
- Lookup table design
- Dashboard layout and chart creation
- Conditional formatting for visual status flags

## File

[`Norman_Inventory_Dashboard.xlsx`](./Norman_Inventory_Dashboard.xlsx)
