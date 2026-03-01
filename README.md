# Excel VBA Sales Data Automation

This project demonstrates automated multi-file data consolidation and reporting using Excel VBA.

## Overview

The solution collects weekly sales data from multiple regional Excel files, processes the data using nested loops and multidimensional arrays, and consolidates the results into a centralized reporting sheet. Pivot tables are refreshed automatically.

## Key Features

- Automated validation of required source files
- Multi-file workbook handling
- Nested loop processing (region × product × week)
- Multidimensional array storage
- Automated report sheet generation
- Pivot table refresh automation
- Basic error handling

## Files

- `sales-data-automation.xlsm` – Excel model with VBA automation
- `data_automation.bas` – Exported VBA module
- `regional-sales-data.zip` – Sample source files

## Purpose

The project simulates a real-world reporting scenario where distributed sales data must be automatically consolidated into a structured management report.
