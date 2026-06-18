# MIS Reporting Automation — Python + Excel

### Reduced manual reporting effort by 60%

---

## Problem
MIS team spending 4 hours daily compiling reports manually
from 5+ warehouse Excel files. Human error in every report.

## Solution
Python scripts automating the full pipeline:
Data ingestion → Cleaning → Transformation → Report generation

## Scripts

| Script | Purpose | Time Saved |
|--------|---------|------------|
| daily_mis_report.py | Auto-generate daily MIS | 2 hrs/day |
| weekly_summary_report.py | Weekly KPI summary | 1 hr/week |
| monthly_dashboard_export.py | Monthly management pack | 4 hrs/month |
| data_cleaning_pipeline.py | Data validation and cleaning | 1 hr/day |
| multi_warehouse_consolidator.py | Merge 5+ warehouse files | 30 min/day |

## Tech Stack
Python | Pandas | OpenPyXL | XlsxWriter | NumPy

## How to Run
pip install -r requirements.txt
python scripts/daily_mis_report.py

## Author
Dilip Patekar | dppatekar12@gmail.com
