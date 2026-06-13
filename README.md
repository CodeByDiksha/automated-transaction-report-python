# Automated Daily Transaction Report — Python & Pandas

## Overview
This project automates the generation of daily banking transaction summary reports — directly inspired by real report automation work at Infosys where daily client transaction reports were manually prepared, taking 30 minutes every day.

This pipeline replaced that manual process completely: it reads raw CSV transaction data, validates it, cleans it, generates summaries, and produces a formatted Excel report — automatically, every day.

## Problem It Solves
- Manual daily report preparation: 30 minutes every day
- Human errors in copy-paste and data entry
- No standardized validation before report generation
- No audit trail of data issues

## Solution
Fully automated Python pipeline that:
1. Reads raw transaction CSV data
2. Runs data quality validation (nulls, duplicates, invalid amounts)
3. Cleans and standardizes the data
4. Generates summary statistics by transaction type and date
5. Produces a formatted multi-sheet Excel report
6. Logs everything for audit trail

**Result: 30 minutes daily manual effort reduced to zero.**

## Tech Stack
- Python 3.x
- Pandas — data ingestion, cleaning, transformation
- NumPy — statistical calculations
- openpyxl — Excel report generation and formatting
- logging — automated audit trail

## Project Structure
```
project2_python/
│
├── transaction_report_automation.py    # Main pipeline script
├── requirements.txt                    # Dependencies
├── data/
│   └── transactions.csv               # Sample input data
├── output/
│   └── transaction_report_YYYY-MM-DD.xlsx   # Generated report
└── logs/
    └── report_YYYY-MM-DD.log          # Execution logs
```

## How to Run

### 1. Install dependencies
```bash
pip install pandas numpy openpyxl
```

### 2. Add your data
Place your transaction CSV file in the `data/` folder as `transactions.csv`

Expected columns: `transaction_id, account_id, amount, transaction_date, transaction_type`

### 3. Run the pipeline
```bash
python transaction_report_automation.py
```

### 4. Check output
- Excel report → `output/transaction_report_YYYY-MM-DD.xlsx`
- Execution log → `logs/report_YYYY-MM-DD.log`

## Excel Report Sheets
| Sheet | Contents |
|---|---|
| Summary | Overall KPIs — total transactions, amounts, unique accounts |
| By Type | Breakdown by transaction type (CREDIT/DEBIT/TRANSFER) |
| By Date | Daily volume and amount trends |
| All Transactions | Full cleaned dataset |
| Flagged Records | Records that failed validation (highlighted in red) |
| Validation Log | List of all data quality issues found |

## Validation Checks Applied
- NULL values in critical fields (transaction_id, account_id, amount, date)
- Duplicate transaction IDs
- Invalid amount values (below minimum or above maximum threshold)
- Invalid transaction types
- Unparseable date formats

## Key Features
- Modular design — each step is a separate function, easy to maintain
- Full exception handling — pipeline logs errors and fails gracefully
- Automated logging — every run creates a timestamped log file
- Formatted Excel output — blue headers, red flagged records, auto-fitted columns
- Configurable — thresholds and paths controlled via CONFIG dictionary

## Author
Diksha Mulik | Data Analyst | [LinkedIn](your-linkedin-url)
