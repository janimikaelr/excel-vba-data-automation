# Excel VBA Sales Data Automation

This project demonstrates automated multi-file data consolidation and reporting using Excel VBA.

## Overview

The solution simulates a real-world reporting scenario where weekly sales data is stored across multiple regional Excel files.

The VBA automation:

- Validates that required source files exist
- Opens each regional workbook dynamically
- Extracts weekly sales data
- Stores the data in a multidimensional array (week × region × product)
- Writes consolidated results into a centralized reporting sheet
- Automatically refreshes pivot tables

The result is a structured and automated reporting workflow that eliminates manual data consolidation.

---

## Technical Implementation

The solution is built using structured VBA modules and includes:

- Multi-file workbook handling (open → read → close)
- Nested loop processing
- Multidimensional array storage for performance
- Data normalization into reporting format
- Automated pivot refresh logic
- Clear modular macro structure

Core workflow:

1. Validate source folder and files  
2. Read sales data into a 3D array  
3. Write consolidated data to master sheet  
4. Refresh reporting pivots  

---

## Technical Skills Demonstrated

- VBA automation and file handling  
- Excel object model manipulation  
- Array-based data processing  
- Nested loop logic  
- Report automation  
- Structured macro design  

---

## Files

- `sales-data-automation.xlsm` – Excel model with VBA automation  
- `data_automation.bas` – Exported VBA module  
- `regional-sales-data.zip` – Sample source files  

---

## Business Context

In many organizations, operational data is distributed across multiple files and locations. This project demonstrates how Excel VBA can be used to automate data consolidation and reporting, improving efficiency and reducing manual workload.
