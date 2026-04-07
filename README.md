README-Vorschlag

Den kannst du fast direkt so in README.md packen:

# Python Excel Report Automation

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

Note: In the current version, configuration paths and credentials are defined directly in the script. For production or public portfolio use, these should be moved to environment variables or a separate config file.

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
- passende Topics.
