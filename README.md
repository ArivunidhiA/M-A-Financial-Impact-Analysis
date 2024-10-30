# M&A Financial Impact Analysis

## Project Overview
This repository contains a comprehensive M&A (Mergers & Acquisitions) financial impact analysis model built in Excel with VBA automation. The model analyzes the potential acquisition impact using historical financial data from public companies.

## Data Source
The model uses historical financial data from:
1. Yahoo Finance (https://finance.yahoo.com/)
   - Historical stock prices
   - Financial statements (Income Statement, Balance Sheet, Cash Flow)
2. SEC EDGAR Database (https://www.sec.gov/edgar/searchedgar/companysearch)
   - 10-K and 10-Q filings for detailed financial information

Sample companies analyzed: 
- Acquirer: Microsoft Corporation (MSFT)
- Target: Activision Blizzard (ATVI)
*Note: You can replace these with any public companies of your choice*

## Repository Contents
1. `MA_Analysis_Model.xlsm` - Main Excel workbook containing:
   - Financial statements
   - Valuation models
   - Synergy analysis
   - Deal accretion/dilution analysis
   - Automated VBA tools

2. `Data_Collection_Scripts.bas` - VBA modules for:
   - Yahoo Finance data import
   - Financial statement formatting
   - Ratio calculations

3. `Documentation/` folder containing:
   - User guide
   - Model assumptions
   - Calculation methodologies

## Key Features
1. **Financial Statement Analysis**
   - 5-year historical financials
   - Key ratio analysis
   - Growth trend analysis

2. **Valuation Methods**
   - Comparable Company Analysis
   - Precedent Transaction Analysis
   - DCF Valuation
   - LBO Analysis

3. **Deal Impact Analysis**
   - EPS Accretion/Dilution
   - Balance Sheet Impact
   - Cash Flow Impact
   - Synergy Modeling

4. **Automated Tools**
   - Data import automation
   - Financial statement standardization
   - Ratio calculations
   - Sensitivity analysis

## Setup Instructions
1. Download all files from the repository
2. Enable macros in Excel
3. Update company tickers in the "Input" sheet if analyzing different companies
4. Run the "Refresh_Data" macro to import latest financial data
5. Review and adjust assumptions in the "Assumptions" sheet

## Usage
1. Input company identifiers and deal assumptions
2. Run data collection macros
3. Review automated analysis
4. Adjust synergy assumptions
5. Generate summary outputs

## Requirements
- Microsoft Excel 2016 or later
- VBA enabled
- Internet connection for data updates
