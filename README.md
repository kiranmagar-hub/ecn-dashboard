# BEF ECN Metrics Generator

**Version:** 2.0  
**Date:** July 2, 2026  
**Author:** Kevin Magar

## Overview

The BEF ECN Metrics Generator is a powerful GUI application that analyzes ECN (Engineering Change Notice) data from Excel files and generates comprehensive metrics reports.

### What It Generates

- **Interactive HTML Dashboard** with charts and KPIs
- **PowerPoint Presentation** with executive summary  
- **Excel Reports** for Void and Hold ECNs
- **JSON Data Export** for further analysis

## Quick Start Guide

### 1. Install Python & Libraries

```bash
# Install Python from: https://www.python.org/downloads/
# Then install required libraries:
pip install pandas openpyxl python-pptx pyodbc
```

**Note:** `pyodbc` is only needed for SQL Server database connections. If using Excel files only, you can skip it.

### 2. Run the Generator

Double-click `ECN_Metrics_Generator.py` or run:
```bash
python ECN_Metrics_Generator.py
```

### 3. Select Data Source

**Option A: Excel File**
- Must have sheet named `Document_TB11`
- Required columns: RequestNum, SubmitDate, ProcCT(days), TotalCT(Days), State, EcnTopic, ECNCoordinator

**Option B: SQL Server Database** (New!)
- Server: wilmatom1f
- Database: WWECNRequests
- Table: Document_TB11
- Click "Test Connection" to verify

### 4. Choose Date Range

- **Default:** Last 365 days from today
- **Custom:** Enter start/end dates (YYYY-MM-DD)
- **All Data:** Click "Clear Dates"

### 5. Generate Reports

Click "Generate Reports" and wait for completion!

## Key Features

✅ **SQL Server Database Support** - Pull data directly from database (NEW!)  
✅ **Excel File Support** - Traditional Excel file import  
✅ **ECN Type 9 Filtering** - Automatically excludes restricted ECNs  
✅ **365-Day Default Range** - Auto-set to last year  
✅ **Coordinator Mapping** - Converts usernames to real names (if ECN Coordinators.xlsx provided)  
✅ **Comprehensive KPIs** - Void rate, FTR, hold rate, percentiles  
✅ **Interactive Charts** - All metrics visualized  

## Output Files

Each run creates a date-stamped folder containing:

- `dashboard.html` - Interactive web dashboard (open in browser)
- `BEF_ECN_Metrics.pptx` - PowerPoint presentation
- `ECN_Void_by_Reason.xlsx` - Void ECNs by reason
- `ECN_Hold_by_Reason.xlsx` - Hold ECNs by reason  
- `ECN_90th_Percentile.xlsx` - Slowest 10% ECNs
- `data.json` - All metrics data
- `source_data.xlsx` - Copy of input file
- `SUMMARY.txt` - Text summary

## Important Notes

### ECN Type 9 Excluded

All ECN Type 9 (Restricted ECNs) are automatically excluded from metrics. This includes:
- `9`
- `(9)` 
- `(9) Restricted ECN`
- `9 - Restricted`

### Coordinator Name Mapping (Optional)

Create `ECN Coordinators.xlsx` in the same folder with columns:
- `Username` - User ID
- `Coordinator` - Full name

Example:
```
Username | Coordinator
---------|-------------
JSMITH   | John Smith
KJONES   | Karen Jones
```

## Troubleshooting

**"Module not found"**  
→ Run: `pip install pandas openpyxl python-pptx`

**"Sheet Document_TB11 not found"**  
→ Your Excel must have this exact sheet name

**No charts in dashboard**  
→ Regenerate reports (old dashboard has old data)

**Encoding errors**  
→ Fixed in this version (UTF-8)

## System Requirements

- Windows 10/11, macOS, or Linux
- Python 3.8 or higher
- 4GB RAM minimum (8GB for large files)
- 500MB disk space

## Files Included

**Required:**
- ECN_Metrics_Generator.py (main application)
- process_data.py (data processor)
- create_enhanced_dashboard.py (HTML generator)
- create_powerpoint.py (PowerPoint generator)
- create_void_report.py (Void report)
- create_hold_report.py (Hold report)
- ADI-Logo-RGB-FullColor.png (logo)

**Optional:**
- ECN Coordinators.xlsx (username mapping)

## Support

**Contact:** kmagar@analog.com  
**Location:** BEF Engineering

---

**Version 2.0** - July 2, 2026  
Internal use only - Analog Devices, Inc.
