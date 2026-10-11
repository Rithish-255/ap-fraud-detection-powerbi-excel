# Power BI Dashboard Step-by-Step Implementation Plan
## Accounts Payable & Duplicate Payment Detection

---

## Table of Contents
1. Prerequisites and Setup
2. Data Preparation
3. Data Model Design
4. DAX Measures and Calculations
5. Dashboard Page 1: Executive Overview
6. Dashboard Page 2: Vendor Risk Analysis
7. Dashboard Page 3: Payment Exception Analysis
8. Dashboard Page 4: Control Review Dashboard
9. Formatting and Optimization
10. Publishing and Deployment

---

## STEP 1: PREREQUISITES AND SETUP

### 1.1 Software Requirements
- Microsoft Power BI Desktop (latest version)
- Microsoft Excel (for data preparation)
- Sample AP Transaction CSV or Excel file

### 1.2 Installation
1. Download Power BI Desktop from: https://powerbi.microsoft.com/en-us/desktop/
2. Install the application
3. Launch Power BI Desktop
4. Create a new blank report

### 1.3 Project Structure
- Data source: AP_Transactions.csv or Excel file
- Model name: AP_Fraud_Detection_Model
- Report name: AP_Risk_Dashboard
- Expected file size: ~5-10 MB

---

## STEP 2: DATA PREPARATION

### 2.1 Prepare Excel/CSV Source File
Create a file named: AP_Transactions.csv

Column Headers:
```
Invoice_ID,Vendor_Name,Vendor_ID,Invoice_Date,Due_Date,Payment_Date,Amount,Currency,Payment_Mode,Department,Approval_Status,Approved_By,PO_Number,Payment_Reference,Duplicate_Flag,Rogue_Payment_Flag,Days_To_Pay,Risk_Score
```

### 2.2 Sample Data Structure
```
INV-1001,Vendor_01,V01,2025-01-02,2025-01-12,2025-01-11,12000,INR,ACH,Procurement,Approved,MAN-04,PO-2001,REF-1001,No,No,9,25
INV-1002,Vendor_02,V02,2025-01-03,2025-01-13,2025-01-15,8500,INR,Wire,Finance,Approved,MAN-07,PO-2002,REF-1002,No,No,12,18
INV-1003,Vendor_03,V03,2025-01-04,2025-01-14,2025-01-16,3200,INR,ACH,Operations,Approved,MAN-05,PO-2003,REF-1003,No,No,12,16
INV-1004,Vendor_01,V01,2025-01-05,2025-01-15,2025-01-18,12000,INR,ACH,Procurement,Approved,MAN-04,PO-2004,REF-1001,Yes,No,13,95
INV-1005,Vendor_08,V08,2025-01-08,2025-01-18,2025-01-10,28000,INR,Wire,AP,Rejected,NA,PO-2005,REF-1005,No,Yes,2,88
```

### 2.3 Data Cleaning in Power BI
1. Open Power BI Desktop
2. Click: Home > Get Data > Text/CSV
3. Select AP_Transactions.csv file
4. Click Load
5. Power Query window opens
6. Review data types:
   - Invoice_ID: Text
   - Vendor_Name: Text
   - Vendor_ID: Text
   - Invoice_Date: Date
   - Due_Date: Date
   - Payment_Date: Date
   - Amount: Decimal Number
   - Currency: Text
   - Payment_Mode: Text
   - Department: Text
   - Approval_Status: Text
   - Approved_By: Text
   - PO_Number: Text
   - Payment_Reference: Text
   - Duplicate_Flag: Text
   - Rogue_Payment_Flag: Text
   - Days_To_Pay: Whole Number
   - Risk_Score: Whole Number
7. Click "Close & Apply"

### 2.4 Handle Missing Values
1. In Power Query, identify null values
2. For approval-related fields: replace with "Missing"
3. For numeric fields: replace with 0
4. Apply transformations and close

---

## STEP 3: DATA MODEL DESIGN

### 3.1 Create Dimension Tables

#### Step 3.1a: Vendor Dimension
1. Right-click in Model View
2. New Table
3. Enter DAX:
```
Vendor_Dimension = DISTINCT(AP_Transactions[Vendor_ID])
```
4. Add columns:
   - Vendor_ID
   - Vendor_Name
   - Vendor_Risk_Level (calculated)

#### Step 3.1b: Department Dimension
1. New Table
2. DAX:
```
Department_Dimension = DISTINCT(AP_Transactions[Department])
```

#### Step 3.1c: Approval Status Dimension
1. New Table
2. DAX:
```
Approval_Status_Dimension = DISTINCT(AP_Transactions[Approval_Status])
```

#### Step 3.1d: Payment Mode Dimension
1. New Table
2. DAX:
```
Payment_Mode_Dimension = DISTINCT(AP_Transactions[Payment_Mode])
```

### 3.2 Create Relationships
1. Go to Model View (left sidebar)
2. Drag lines to create relationships:
   - AP_Transactions[Vendor_ID] → Vendor_Dimension[Vendor_ID] (One-to-Many)
   - AP_Transactions[Department] → Department_Dimension[Department] (One-to-Many)
   - AP_Transactions[Approval_Status] → Approval_Status_Dimension[Approval_Status] (One-to-Many)
   - AP_Transactions[Payment_Mode] → Payment_Mode_Dimension[Payment_Mode] (One-to-Many)
3. Verify relationship type: Many-to-One

### 3.3 Mark as Date Table
1. Right-click on AP_Transactions table
2. Select "Mark as Date Table"
3. Column: Invoice_Date
4. Click OK

---

## STEP 4: DAX MEASURES AND CALCULATIONS

### 4.1 Basic Count Measures

#### Measure 1: Total Invoices
1. Right-click AP_Transactions table > New Measure
2. Name: Total Invoices
3. Enter DAX:
```
Total Invoices = DISTINCTCOUNT(AP_Transactions[Invoice_ID])
```
4. Format: Whole Number
5. Press Enter

#### Measure 2: Total Payment Amount
1. New Measure
2. Name: Total Amount
3. DAX:
```
Total Amount = SUM(AP_Transactions[Amount])
```
4. Format: Currency (INR)

#### Measure 3: Total Unique Vendors
1. New Measure
2. Name: Total Vendors
3. DAX:
```
Total Vendors = DISTINCTCOUNT(AP_Transactions[Vendor_ID])
```
4. Format: Whole Number

### 4.2 Duplicate Payment Measures

#### Measure 4: Duplicate Payment Count
1. New Measure
2. Name: Duplicate Payment Count
3. DAX:
```
Duplicate Payment Count = 
CALCULATE(
    COUNTROWS(AP_Transactions),
    AP_Transactions[Duplicate_Flag] = "Yes"
)
```
4. Format: Whole Number

#### Measure 5: Duplicate Payment Amount
1. New Measure
2. Name: Duplicate Payment Amount
3. DAX:
```
Duplicate Payment Amount = 
CALCULATE(
    SUM(AP_Transactions[Amount]),
    AP_Transactions[Duplicate_Flag] = "Yes"
)
```
4. Format: Currency (INR)

#### Measure 6: Duplicate Payment %
1. New Measure
2. Name: Duplicate Payment %
3. DAX:
```
Duplicate Payment % = 
DIVIDE(
    [Duplicate Payment Count],
    [Total Invoices],
    0
) * 100
```
4. Format: Decimal (2 decimals)

### 4.3 Rogue Payment Measures

#### Measure 7: Rogue Payment Count
1. New Measure
2. Name: Rogue Payment Count
3. DAX:
```
Rogue Payment Count = 
CALCULATE(
    COUNTROWS(AP_Transactions),
    AP_Transactions[Rogue_Payment_Flag] = "Yes"
)
```
4. Format: Whole Number

#### Measure 8: Rogue Payment Amount
1. New Measure
2. Name: Rogue Payment Amount
3. DAX:
```
Rogue Payment Amount = 
CALCULATE(
    SUM(AP_Transactions[Amount]),
    AP_Transactions[Rogue_Payment_Flag] = "Yes"
)
```
4. Format: Currency (INR)

### 4.4 Approval Measures

#### Measure 9: Approval Bypass Count
1. New Measure
2. Name: Approval Bypass Count
3. DAX:
```
Approval Bypass Count = 
CALCULATE(
    COUNTROWS(AP_Transactions),
    OR(
        AP_Transactions[Approved_By] = "NA",
        ISBLANK(AP_Transactions[Approved_By])
    )
)
```
4. Format: Whole Number

#### Measure 10: Approved Invoices
1. New Measure
2. Name: Approved Invoices
3. DAX:
```
Approved Invoices = 
CALCULATE(
    COUNTROWS(AP_Transactions),
    AP_Transactions[Approval_Status] = "Approved"
)
```
4. Format: Whole Number

#### Measure 11: Rejected Invoices
1. New Measure
2. Name: Rejected Invoices
3. DAX:
```
Rejected Invoices = 
CALCULATE(
    COUNTROWS(AP_Transactions),
    AP_Transactions[Approval_Status] = "Rejected"
)
```
4. Format: Whole Number

### 4.5 Payment Timing Measures

#### Measure 12: Avg Days To Pay
1. New Measure
2. Name: Avg Days To Pay
3. DAX:
```
Avg Days To Pay = AVERAGE(AP_Transactions[Days_To_Pay])
```
4. Format: Decimal (1 decimal)

#### Measure 13: Max Days To Pay
1. New Measure
2. Name: Max Days To Pay
3. DAX:
```
Max Days To Pay = MAX(AP_Transactions[Days_To_Pay])
```
4. Format: Whole Number

#### Measure 14: Delayed Payments
1. New Measure
2. Name: Delayed Payments
3. DAX:
```
Delayed Payments = 
CALCULATE(
    COUNTROWS(AP_Transactions),
    AP_Transactions[Days_To_Pay] > 15
)
```
4. Format: Whole Number

### 4.6 Risk Score Measures

#### Measure 15: Avg Risk Score
1. New Measure
2. Name: Avg Risk Score
3. DAX:
```
Avg Risk Score = AVERAGE(AP_Transactions[Risk_Score])
```
4. Format: Decimal (2 decimals)

#### Measure 16: High Risk Invoices
1. New Measure
2. Name: High Risk Invoices
3. DAX:
```
High Risk Invoices = 
CALCULATE(
    COUNTROWS(AP_Transactions),
    AP_Transactions[Risk_Score] >= 75
)
```
4. Format: Whole Number

#### Measure 17: Critical Risk Invoices
1. New Measure
2. Name: Critical Risk Invoices
3. DAX:
```
Critical Risk Invoices = 
CALCULATE(
    COUNTROWS(AP_Transactions),
    AP_Transactions[Risk_Score] >= 90
)
```
4. Format: Whole Number

### 4.7 Department-Level Measures

#### Measure 18: Exception Rate by Department
1. New Measure
2. Name: Exception Rate %
3. DAX:
```
Exception Rate % = 
DIVIDE(
    [Duplicate Payment Count] + [Rogue Payment Count],
    [Total Invoices],
    0
) * 100
```
4. Format: Decimal (2 decimals)

---

## STEP 5: DASHBOARD PAGE 1 - EXECUTIVE OVERVIEW

### 5.1 Create New Page
1. Right-click on page tab at bottom
2. Select "New Page"
3. Rename to: "Executive Overview"

### 5.2 Add Page Background
1. Click on page background
2. Format > Page Background > Color: Light Gray

### 5.3 Add Title
1. Insert > Text Box
2. Enter: "Accounts Payable Risk Dashboard - Executive Overview"
3. Font: 24pt, Bold, Dark Blue
4. Position at top

### 5.4 Add KPI Cards (Row 1)
Create 5 cards in the first row:

#### Card 1: Total Invoices
1. Insert > Card (from Visualizations)
2. Drag "Total Invoices" measure to Values
3. Title: "Total Invoices"
4. Format card: Blue background, white text
5. Position: Top-left

#### Card 2: Total Payment Value
1. Insert > Card
2. Drag "Total Amount" measure to Values
3. Title: "Total Payment Value"
4. Format: Blue background
5. Position: Next to Card 1

#### Card 3: Duplicate Payments
1. Insert > Card
2. Drag "Duplicate Payment Count" measure to Values
3. Title: "Duplicate Payments"
4. Format: Red background (warning color)
5. Position: Next to Card 2

#### Card 4: Rogue Payments
1. Insert > Card
2. Drag "Rogue Payment Count" measure to Values
3. Title: "Rogue Payments"
4. Format: Red background
5. Position: Next to Card 3

#### Card 5: Approval Bypass
1. Insert > Card
2. Drag "Approval Bypass Count" measure to Values
3. Title: "Approval Bypass Cases"
4. Format: Orange background (caution color)
5. Position: Next to Card 4

### 5.5 Add KPI Cards (Row 2)

#### Card 6: Average Days to Pay
1. Insert > Card
2. Drag "Avg Days To Pay" measure
3. Title: "Avg Days to Pay"
4. Format: Gray background
5. Position: Row 2, Left

#### Card 7: High Risk Invoices
1. Insert > Card
2. Drag "High Risk Invoices" measure
3. Title: "High Risk Invoices"
4. Format: Red background
5. Position: Row 2, Next to Card 6

#### Card 8: Total Vendors
1. Insert > Card
2. Drag "Total Vendors" measure
3. Title: "Total Vendors"
4. Format: Blue background
5. Position: Row 2, Next to Card 7

### 5.6 Add Summary Charts

#### Chart 1: Approval Status Distribution
1. Insert > Pie Chart
2. Legend: Approval_Status
3. Values: Total Invoices measure
4. Title: "Approval Status Distribution"
5. Position: Bottom-left
6. Size: 500x300px

#### Chart 2: Exception Types
1. Insert > Stacked Column Chart
2. Axis: "Exception Type" (create calculated column)
3. Value: Count of invoices
4. Title: "Exception Distribution"
5. Position: Bottom-middle
6. Size: 500x300px

#### Chart 3: Payment Amount by Risk
1. Insert > Stacked Bar Chart
2. Axis: Risk_Score buckets
3. Value: Total Amount
4. Title: "Payment Amount by Risk Level"
5. Position: Bottom-right
6. Size: 500x300px

---

## STEP 6: DASHBOARD PAGE 2 - VENDOR RISK ANALYSIS

### 6.1 Create New Page
1. Right-click on page tab
2. New Page
3. Rename to: "Vendor Risk Analysis"

### 6.2 Add Page Title
1. Insert > Text Box
2. Enter: "Vendor Risk Analysis"
3. Format: 20pt, Bold, Dark Blue

### 6.3 Add Filter Slicers

#### Slicer 1: Vendor Filter
1. Insert > Slicer
2. Field: Vendor_Name
3. Format: Dropdown list
4. Position: Top-left
5. Title: "Select Vendor(s)"

#### Slicer 2: Department Filter
1. Insert > Slicer
2. Field: Department
3. Format: List
4. Position: Top-middle
5. Title: "Department"

### 6.4 Add Vendor Risk Table
1. Insert > Table
2. Columns:
   - Vendor_Name
   - Vendor_ID
   - Duplicate_Flag (count)
   - Rogue_Payment_Flag (count)
   - Avg Risk Score
   - Total Amount
3. Sort by Risk Score (descending)
4. Title: "Top Risk Vendors"
5. Position: Below slicers
6. Size: Full width, 400px height

### 6.5 Add Vendor Charts

#### Chart 1: Duplicate Payments by Vendor
1. Insert > Horizontal Bar Chart
2. Axis: Vendor_Name
3. Value: Count of Duplicate_Flag = "Yes"
4. Title: "Duplicate Payments by Vendor"
5. Sort: Descending
6. Position: Bottom-left
7. Size: 500x300px

#### Chart 2: Vendor Amount Distribution
1. Insert > Bar Chart
2. Axis: Vendor_Name (top 10)
3. Value: Total Amount
4. Title: "Total Payment Amount by Vendor"
5. Position: Bottom-middle
6. Size: 500x300px

#### Chart 3: Rogue Payments by Vendor
1. Insert > Treemap
2. Group: Vendor_Name
3. Values: Count of Rogue_Payment_Flag = "Yes"
4. Color: Risk Score
5. Title: "Rogue Payments by Vendor"
6. Position: Bottom-right
7. Size: 500x300px

### 6.6 Add Cross-Filter Interactions
1. Select the Vendor_Name slicer
2. Right-click > Edit Interactions
3. Set "Filter" interaction with all charts on this page
4. Click outside to apply

---

## STEP 7: DASHBOARD PAGE 3 - PAYMENT EXCEPTION ANALYSIS

### 7.1 Create New Page
1. New Page
2. Rename to: "Payment Exception Analysis"

### 7.2 Add Title
1. Insert > Text Box
2. Enter: "Payment Exception Analysis"
3. Format: 20pt, Bold

### 7.3 Add Filter Slicers

#### Slicer 1: Approval Status
1. Insert > Slicer
2. Field: Approval_Status
3. Position: Top-left

#### Slicer 2: Payment Mode
1. Insert > Slicer
2. Field: Payment_Mode
3. Position: Top-middle

#### Slicer 3: Date Range
1. Insert > Slicer
2. Field: Invoice_Date
3. Type: Between (date range)
4. Position: Top-right

### 7.4 Add Exception Charts

#### Chart 1: Exception Types
1. Insert > Stacked Column Chart
2. Create calculated column:
   - "Duplicate" if Duplicate_Flag = "Yes"
   - "Rogue" if Rogue_Payment_Flag = "Yes"
   - "Delay" if Days_To_Pay > 15
   - "Normal" otherwise
3. Axis: Month(Invoice_Date)
4. Values: Count
5. Title: "Exception Trend Over Time"
6. Position: Top-left
7. Size: 600x300px

#### Chart 2: Department-wise Exception Rate
1. Insert > Clustered Bar Chart
2. Axis: Department
3. Value: Exception Rate %
4. Title: "Exception Rate by Department"
5. Position: Top-right
6. Size: 500x300px

#### Chart 3: Payment Mode Risk
1. Insert > Pie Chart
2. Legend: Payment_Mode
3. Values: Count of Rogue_Payment_Flag = "Yes"
4. Title: "Rogue Payments by Payment Mode"
5. Position: Middle-left
6. Size: 400x300px

#### Chart 4: Days to Pay Distribution
1. Insert > Histogram
2. Field: Days_To_Pay
3. Bins: 10
4. Title: "Distribution of Days to Pay"
5. Position: Middle-middle
6. Size: 400x300px

#### Chart 5: Exception Detail Table
1. Insert > Table
2. Columns:
   - Invoice_ID
   - Vendor_Name
   - Amount
   - Approval_Status
   - Duplicate_Flag
   - Rogue_Payment_Flag
   - Days_To_Pay
   - Risk_Score
3. Filter: Show only exceptions (Duplicate_Flag = "Yes" OR Rogue_Payment_Flag = "Yes")
4. Title: "Flagged Exception Invoices"
5. Position: Bottom
6. Size: Full width, 250px height

---

## STEP 8: DASHBOARD PAGE 4 - CONTROL REVIEW DASHBOARD

### 8.1 Create New Page
1. New Page
2. Rename to: "Control Review"

### 8.2 Add Title and Summary
1. Insert > Text Box
2. Enter: "Approval Control and Risk Review"
3. Format: 20pt, Bold

### 8.3 Add Control KPI Cards

#### Card 1: Total Approved
1. Insert > Card
2. Measure: Approved Invoices
3. Title: "Total Approved"
4. Format: Green background

#### Card 2: Total Rejected
1. Insert > Card
2. Measure: Rejected Invoices
3. Title: "Total Rejected"
4. Format: Red background

#### Card 3: Pending
1. Insert > Card
2. Measure: Create new measure for Pending status
3. Title: "Pending"
4. Format: Yellow background

#### Card 4: Compliance Rate
1. Insert > Card
2. Create measure: (Total Invoices - Approval Bypass) / Total Invoices * 100
3. Title: "Approval Compliance %"
4. Format: Blue background

### 8.4 Add Risk Score Charts

#### Chart 1: Risk Score Distribution
1. Insert > Histogram
2. Field: Risk_Score
3. Title: "Risk Score Distribution"
4. Position: Top-left
5. Size: 500x300px

#### Chart 2: Critical Risk Cases
1. Insert > Scatter Chart
2. X-Axis: Days_To_Pay
3. Y-Axis: Amount
4. Size: Risk_Score
5. Color: Approval_Status
6. Title: "Critical Risk Cases (Amount vs Delay)"
7. Position: Top-right
8. Size: 500x300px

#### Chart 3: Approval Chain Compliance
1. Insert > Stacked Column Chart
2. Axis: Department
3. Values: 
   - Count of Approved Invoices (green)
   - Count of Rejected Invoices (red)
   - Count of Pending Invoices (yellow)
4. Title: "Approval Status by Department"
5. Position: Bottom-left
6. Size: 600x300px

#### Chart 4: High Risk Invoice Table
1. Insert > Table
2. Columns:
   - Invoice_ID
   - Vendor_Name
   - Amount
   - Risk_Score
   - Days_To_Pay
   - Approval_Status
   - Approved_By
3. Filter: Risk_Score >= 75
4. Sort: Risk_Score (descending)
5. Title: "High Risk Investigation List"
6. Position: Bottom-right
7. Size: 600x300px

### 8.5 Add Gauge Charts

#### Gauge 1: Exception Rate
1. Insert > Gauge
2. Value: Exception Rate % measure
3. Min: 0, Max: 10
4. Target: 5
5. Title: "Overall Exception Rate %"
6. Position: Middle section

#### Gauge 2: Compliance Score
1. Insert > Gauge
2. Value: Compliance Rate measure (0-100)
3. Min: 0, Max: 100
4. Target: 95
5. Title: "Control Compliance Score"
6. Position: Next to Gauge 1

---

## STEP 9: FORMATTING AND OPTIMIZATION

### 9.1 Apply Consistent Theme
1. View > Themes
2. Select "Blue Tones" or "Modern"
3. All pages update automatically

### 9.2 Add Bookmarks (Interactive Navigation)
1. View > Bookmarks pane
2. Create bookmark: "Executive Overview"
3. Create bookmark: "Vendor Analysis"
4. Create bookmark: "Exception Analysis"
5. Create bookmark: "Control Review"

### 9.3 Add Buttons for Navigation
1. Insert > Buttons > Blank Button
2. Set Action: Bookmark > Executive Overview
3. Repeat for all bookmarks
4. Position in a navigation bar at top of each page

### 9.4 Apply Conditional Formatting
1. Select Risk_Score column in tables
2. Format > Conditional Formatting > Background Color
3. Min: Green (0), Mid: Yellow (50), Max: Red (100)

### 9.5 Optimize Performance
1. Remove unused columns from model
2. Hide intermediate calculated columns
3. Create summary tables for large datasets
4. Reduce visual density on each page

### 9.6 Add Tooltips
1. Select each visual
2. Format > Tooltips > Add custom tooltips
3. Example: "Hover to see vendor details"

---

## STEP 10: PUBLISHING AND DEPLOYMENT

### 10.1 Save the Report
1. File > Save
2. Name: AP_Risk_Dashboard.pbix
3. Location: [Your Documents folder]

### 10.2 Publish to Power BI Service (Optional)
1. File > Publish
2. Select workspace (or create new)
3. Click "Select"
4. Wait for publish confirmation
5. Dashboard is now available on: https://app.powerbi.com

### 10.3 Set Up Data Refresh
1. In Power BI Service
2. Dataset > Settings
3. Data source credentials: Enter CSV file location
4. Refresh schedule: Daily at 8 AM

### 10.4 Share Dashboard
1. Share button (top-right)
2. Add email addresses
3. Set permissions: View only
4. Send invitation

### 10.5 Create Mobile Layout
1. View > Mobile Layout
2. Rearrange visuals for phone screen
3. Remove cluttered elements
4. Test on Power BI mobile app

---

## STEP 11: FINAL CHECKLIST

- [ ] All data loaded and cleaned
- [ ] Data model relationships defined
- [ ] All DAX measures created and tested
- [ ] Executive Overview page complete
- [ ] Vendor Risk Analysis page complete
- [ ] Exception Analysis page complete
- [ ] Control Review page complete
- [ ] Slicers working on all pages
- [ ] Cross-filters configured
- [ ] Formatting applied consistently
- [ ] Bookmarks and buttons created
- [ ] Mobile layout tested
- [ ] Report saved
- [ ] Published to Power BI Service (if applicable)
- [ ] Shared with stakeholders

---

## STEP 12: USER GUIDE

### 12.1 How to Use the Dashboard

**For Executives:**
1. Start with Executive Overview page
2. Check total exception counts (red cards)
3. Review approval status pie chart
4. Use date filter for time period analysis

**For Audit Teams:**
1. Go to Vendor Risk Analysis
2. Select high-risk vendors
3. Review duplicate and rogue payment counts
4. Export the table for detailed audit trail

**For AP Managers:**
1. Use Exception Analysis page
2. Filter by department
3. Review payment timeline trends
4. Check days-to-pay distribution

**For Investigators:**
1. Navigate to Control Review page
2. Use High Risk Investigation List
3. Sort by Risk Score
4. Click invoice details for further review

### 12.2 Interpretation Guide

- **Red cards** = Immediate action required
- **Yellow/Orange** = Watch closely, may escalate
- **Green** = Normal, low risk
- **Risk Score >= 90** = Critical, escalate immediately
- **Exception Rate > 5%** = Control gap, requires review

---

## Conclusion

This step-by-step guide covers the complete development of a Power BI dashboard for AP fraud detection. Follow each step sequentially for best results. Adjust visualizations and measures based on your organization's specific requirements and data structure.

**Expected Development Time:** 4-6 hours
**Skill Level Required:** Intermediate Power BI knowledge
**Maintenance:** Weekly data refresh, monthly review of high-risk vendors
