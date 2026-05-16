# invoice-automation-bot

## Overview
RPA bot that automatically extracts data from PDF invoices 
and saves to Excel using UiPath and Python.

## Tools Used
- UiPath Studio Community Edition
- Python 3.13.2 (PyCharm IDE)
- pdfplumber
- openpyxl
- Windows Task Scheduler

## How It Works
1. Bot monitors input folder for PDF invoices
2. Python extracts Invoice Number, Vendor, Date, Amount
3. Data saved to Excel automatically
4. Scheduled via Windows Task Scheduler to run daily at 9AM

## Project Structure
- Main.xaml — UiPath workflow
- extract.py — Python extraction script
- input/ — PDF invoices folder
- output/ — Generated Excel reports
