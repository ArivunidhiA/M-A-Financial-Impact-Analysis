# Excel Workbook Structure (MA_Analysis_Model.xlsm)

## 1. Input Dashboard
```excel
A1: Deal Parameters
A2: Acquirer Ticker: [MSFT]
A3: Target Ticker: [ATVI]
A4: Deal Premium (%): [45%]
A5: Payment Method: [Cash/Stock/Mixed]
A6: Synergy Assumptions: [$B]

B1: Market Data
B2: Current Stock Price
B3: Shares Outstanding
B4: Market Capitalization
B5: Enterprise Value
```

## 2. Financial Statements
### Income Statement (Both Companies)
```excel
A1: Revenue
A2: Cost of Revenue
A3: Gross Profit
A4: Operating Expenses
A5: EBIT
A6: Interest Expense
A7: Tax
A8: Net Income

Formulas:
A3 = A1 - A2
A5 = A3 - A4
A8 = A5 - A6 - A7
```

### Balance Sheet (Both Companies)
```excel
A1: Assets
A2: Cash & Equivalents
A3: Accounts Receivable
A4: Inventory
A5: Total Current Assets
A6: PP&E
A7: Goodwill
A8: Total Assets

B1: Liabilities & Equity
B2: Accounts Payable
B3: Short-term Debt
B4: Total Current Liabilities
B5: Long-term Debt
B6: Total Liabilities
B7: Shareholders' Equity
B8: Total Liabilities & Equity

Formulas:
A5 = SUM(A2:A4)
A8 = A5 + A6 + A7
B4 = SUM(B2:B3)
B6 = B4 + B5
B8 = B6 + B7
```

## 3. Valuation Analysis
```excel
A1: Trading Multiples
A2: P/E Ratio
A3: EV/EBITDA
A4: EV/Revenue

B1: Transaction Value
B2: Equity Value
B3: Enterprise Value
B4: Premium Paid

Formulas:
A2 = Stock_Price / EPS
A3 = Enterprise_Value / EBITDA
A4 = Enterprise_Value / Revenue
B2 = Shares_Outstanding * Offer_Price
B3 = B2 + Net_Debt
B4 = (Offer_Price / Current_Price - 1) * 100
```

## 4. Deal Impact
```excel
A1: Accretion/Dilution Analysis
A2: Pro Forma EPS
A3: Standalone EPS
A4: % Impact

B1: Balance Sheet Impact
B2: New Debt
B3: New Equity
B4: Goodwill Created

Formulas:
A2 = Combined_Net_Income / New_Shares
A4 = (A2/A3 - 1) * 100
B4 = Purchase_Price - Net_Assets_Acquired
```

## 5. Sensitivity Analysis
```excel
Matrix showing EPS impact based on:
- Deal Premium (columns)
- Synergy Achievement (rows)
- Cost of Debt (tabs)

Formulas use Data Table function for automated calculation
```

## 6. Output Summary
```excel
Key Metrics Dashboard showing:
- Deal Overview
- Strategic Rationale
- Financial Impact
- Risk Factors
- Recommendation
```
