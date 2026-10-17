# SQL to Excel Report

Runs SQL queries on a database and writes a formatted Excel report, one sheet per query.

## What it does
- Builds a small sample SQLite database (customers, products, orders)
- Runs 3 queries using `JOIN`, `GROUP BY` and `ORDER BY`
- Writes `sales_report.xlsx` with: Sales by Product, Sales by Month, Top Customers
- Styled header row, money formatting, auto column widths, frozen header

## Setup
```
python -m venv venv
venv\Scripts\activate        # Mac/Linux: source venv/bin/activate
pip install -r requirements.txt
```

## Run
```
python report.py
python report.py --db my_database.db --output report.xlsx
```

## Adapting to a client
Edit the `QUERIES` dictionary. Each entry becomes one Excel sheet. To use MySQL or PostgreSQL, replace `sqlite3.connect()` with that database's connector.

See `sample_output/sales_report.xlsx` for an example result.
