# warehouse-reporting-automation

Automated reporting pipeline for warehouse/operations data using Python, SQL, Excel export, historical KPI tracking, and Outlook integration.

## Overview

This project automates the daily creation of an operational Excel report for MOK-related warehouse positions.

The pipeline:

1. queries raw position data from a database via ODBC
2. normalizes and enriches the dataset with business-specific calculations
3. builds KPI tables for reporting
4. updates historical KPI records
5. exports the final report to Excel
6. optionally sends the report via Outlook

The goal is to reduce manual reporting effort and provide a reproducible daily reporting workflow.

## Business Context

The script is designed for operational reporting in a warehouse/logistics environment.

It focuses on:
- daily position data
- volume-related KPIs
- picks per position
- container and stop-point metrics
- hourly KPI aggregation
- comparison with historical performance

## Features

- ODBC database connection
- SQL-based extraction of operational data
- business-specific KPI calculations
- lookup-based enrichment (e.g. route / shelf-meter mapping)
- daily MOK summary generation
- hourly KPI matrix
- historical KPI storage and deduplication
- Excel export with multiple report sheets
- logging to file and console
- Outlook mail integration

## Tech Stack

- Python
- pandas
- pyodbc
- openpyxl
- win32com / Outlook COM
- Excel
- SQL / ODBC

## Project Structure

```text
.
├── main.py
├── README.md
└── output/
Workflow
1. Extract

The script queries operational position data for a selected business date from the source database.

2. Transform

The raw dataset is cleaned and enriched with additional calculated fields such as:

volume per quantity
picks per position
unpacked share
route category / shelf-meter mapping
3. Report

Three main output tables are generated:

MOK summary KPIs
hourly KPI breakdown
comparison block against historical values
4. History

The KPI result set is appended to the historical dataset and deduplicated using a generated HIST_ID.

5. Export

The final report is written to an Excel file.

Example KPIs

The report includes metrics such as:

MOK positions
volume per position
picks per position
positions per container
positions per stop-point
volume utilization
share of unpacked positions
hourly KPI distribution
Configuration

The script currently uses hardcoded configuration values for:

ODBC DSN
database credentials
export paths
lookup file paths
mail recipients

Example values:

DB_DSN = "YOUR_DSN_NAME"
DB_USER = "YOUR_USERNAME"
DB_PASSWORD = "YOUR_PASSWORD"
EXPORT_DIR = r"\\server\\share\\folder\\exports"

For a production-ready version, use:

.env
environment variables
secret management
separated config modules
Requirements

Install dependencies:

pip install pandas pyodbc openpyxl pywin32
How to Run
python main.py
Output

The script generates:

an Excel report file
a log file
an updated history file
Limitations

This repository is a portfolio version of an internal reporting workflow.

Because of the original business environment, some parts depend on:

internal database access
company-specific network paths
Outlook on Windows
internal lookup files

To keep the repository shareable, sensitive paths, recipients, and credentials should be replaced with placeholders or environment-based configuration.

Possible Improvements
move configuration into .env
add parquet support for history files or align file extension with current implementation
split the script into multiple modules
add schema validation for inputs
improve exception handling
provide sample input/output data
add screenshots of the generated report
add automated tests for KPI calculations
Why this project matters

This project demonstrates practical skills in:

process automation
data extraction from operational systems
business-oriented KPI engineering
Excel reporting
Python in a real warehouse/logistics use case

It is not a tutorial project, but a real reporting automation workflow adapted into a portfolio-friendly format.
