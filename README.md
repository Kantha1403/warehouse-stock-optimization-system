# Warehouse Stock Optimization System

## Overview

Built an automated warehouse inventory optimization and reorder planning system integrating multiple ERPNext APIs across 16 service warehouses.

The system automated spare-parts analysis, reduced manual consolidation effort, and generated SKU-level shortage/excess insights through scheduled Excel reporting.

---

## Business Problem

Warehouse stock planning was previously dependent on manual ERP exports and Excel-based consolidation across multiple service locations.

Challenges included:

- Manual extraction from multiple ERP modules
- Delayed shortage visibility
- No standardized reorder logic
- High effort for monthly consolidation
- Difficulty tracking excess inventory

The manual process required several days per reporting cycle.

---

## Solution

Developed a Python-based automation pipeline that:

- Integrated 6 ERPNext API modules
- Processed 4,300+ spare part SKUs
- Classified parts into:
  - Accessories
  - Breakdown Parts
- Computed reorder requirements using demand-driven logic
- Flagged shortage, excess, and inventory gaps
- Generated automated multi-sheet Excel reports
- Executed automatically using n8n

---

## Reorder Logic

The reorder quantity calculation was based on average demand, lead time, and safety stock multiplier.

Reorder Formula:

Reorder Qty = Average Demand × 45-Day Lead Time × 1.5 Safety Factor

---

## System Workflow

```text
ERPNext APIs
      ↓
Data Extraction
      ↓
SKU Classification
      ↓
Demand Analysis
      ↓
Reorder Calculation
      ↓
Inventory Gap Detection
      ↓
Excel Report Generation
      ↓
Scheduled Automation
```

---

## Key Features

- ERPNext REST API integration
- Automated warehouse stock analysis
- SKU categorization engine
- Demand-based reorder planning
- Shortage and excess inventory detection
- Automated Excel report generation
- Scheduled execution via n8n
- Multi-warehouse inventory visibility

---

## Tech Stack

- Python
- Pandas
- NumPy
- ERPNext REST API
- OpenPyXL
- n8n

---

## Business Impact

- Reduced manual reporting effort from days to minutes
- Improved visibility into warehouse inventory gaps
- Standardized reorder planning methodology
- Enabled proactive spare-parts procurement
- Automated recurring reporting workflow

---

## Disclaimer

This repository is a portfolio representation of work completed during an internship.

Source code, credentials, internal ERP configurations, and business-sensitive logic have been excluded to comply with company confidentiality policies.
