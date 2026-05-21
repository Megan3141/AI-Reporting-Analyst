# AI Reporting Analyst

Python-based reporting and analytics pipeline for extracting, standardising and consolidating semi-structured Excel reporting data.

## Overview

This project automates the ingestion and transformation of historical reporting files stored across complex Excel folder structures.

The pipeline is designed to handle semi-structured reporting workbooks that contain:
- inconsistent layouts
- changing worksheet names
- subtotal/summary rows
- evolving schemas
- nested folder structures
- historical reporting periods

The goal is to convert manually maintained reporting files into clean analytical datasets suitable for:
- reporting automation
- analytics workflows
- dashboarding
- AI-assisted analysis
- downstream data engineering pipelines

---

## Features

- Automated Excel workbook ingestion
- Weekly file detection and filtering
- Dynamic worksheet detection
- Historical season and period extraction
- Semi-structured row parsing
- Removal of totals and invalid rows
- Zero-value filtering
- Consolidated master dataset creation
- Recovery extraction workflows for inconsistent files
- Extraction logging and validation

---

## Technologies Used

- Python
- pandas
- openpyxl

---

## Example Workflow

```text
Excel Workbooks
       ↓
Folder Scanning
       ↓
Worksheet Detection
       ↓
Row Extraction
       ↓
Data Cleaning
       ↓
Validation & Logging
       ↓
Master Dataset Creation
```


## Current Focus

Building scalable reporting pipelines and automation tooling for semi-structured operational data workflows.

---

## Disclaimer

This repository contains generic extraction and transformation logic only.
