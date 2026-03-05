# Excel Payslip Automation System

Automated HR payroll distribution system built with Excel VBA.

## Overview

This project automates payslip generation and distribution for large organizations.

Features include:

- Automatic PDF generation per employee
- Auto folder structure per Business Unit and Outlet
- Duplicate prevention system
- Email distribution list generation
- ZIP packaging per outlet
- Ultra-robust header detection
- Dynamic employee selection using index system

The system can process 300+ employees within seconds.

## Problem Solved

HR teams often manually export payslips one by one.  
This project automates the entire process.

Before:
Manual export → 2-3 hours work

After:
Fully automated → under 1 minute

## Key Features

• Automatic PDF generation  
• Dynamic data detection  
• Payroll index navigation system  
• Outlet-based file grouping  
• Batch ZIP compression  
• Email distribution preparation  
• Enterprise-ready Excel Add-in  

## Tech Stack

- Excel VBA
- File System Automation
- Shell ZIP integration
- HR data processing logic

## System Flow

1. HR opens payroll template
2. Add-in scans employee database
3. System loops employee index
4. Payslip generated automatically
5. PDF exported per employee
6. Files grouped per outlet
7. ZIP generated per outlet
8. Email distribution file created

## Example Output Structure

EXPORT_PDF
 ├─ Unit_Bisnis_A
 │   ├─ Outlet_1
 │   │   ├─ Employee_A.pdf
 │   │   ├─ Employee_B.pdf
 │   │   └─ Outlet_1.zip

## Author

Zenobius Oktavianus Setiawan 
Data Analyst | Automation Enthusiast
