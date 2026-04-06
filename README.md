# Excel Reporting Automation with Python

This project demonstrates a Windows-based reporting workflow using Python, ODBC, Excel COM automation, and Outlook draft creation.

## What it does

- Queries operational data from a database via ODBC
- Writes the result into a source worksheet in an Excel workbook
- Inserts worksheet formulas and converts them to values
- Updates a history sheet
- Removes duplicates from the history sheet
- Exports a formatted report as `.xlsx`
- Creates an Outlook draft email with the exported report attached
- Uses a local-copy workflow to reduce the risk of locking a shared workbook directly

## Tech stack

- Python
- pandas
- pyodbc
- pywin32
- Microsoft Excel Desktop
- Microsoft Outlook Desktop
- Windows environment with ODBC DSN configured

## Why this project is interesting

This is a practical example of process automation in an enterprise environment.  
It combines:

- database access
- Excel automation
- file locking and atomic replacement
- reporting workflows
- Outlook integration

## Setup

### 1. Clone the repository

```bash
git clone <your-repo-url>
cd <your-repo-name>