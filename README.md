# VBA-challenge
Module 2 Challenge - VBA Scripting 

# Quarterly Stock Data Analysis - VBA Scripting

## Description

This VBA script is created to automate the analysis of quarterly stock data within an excel workbook. The script performs the following tasks:

1. **Calculate Quarterly Changes:**
   - Computes the quarterly change for each stock from the opening price at the beginning of the quarter to the closing price at the end of the quarter.
   - Determines the percentage change for each stock over the quarter.
   - Sums the total stock volume traded over the quarter.

2. **Identify Key Stocks:**
   - Finds the stock with the greatest percentage increase.
   - Finds the stock with the greatest percentage decrease.
   - Finds the stock with the greatest total volume.

3. **Output Results:**
   - Outputs the ticker, quarterly change, percentage change, and total volume for each stock in the worksheet.
   - Displays the stock with the greatest percentage increase, greatest percentage decrease, and greatest total volume.

4. **Conditional Formatting:**
   - Applies conditional formatting to the quarterly change column to color code cells: red for negative values and green for positive values.
   - Leaves values of '0' as white

5. **Cell Alignment:**
   - Centers all cells containing the output data for better readability.
  
6. **Button:**
   - A button labeled "Calculate Quarterly Stock Data" runs the script from VBA editor for easier accessability

## How to Use

1. Included is a button in each Quarterly Sheet with the VBA script - Click it!

Alternatively

1. Open the VBA Editor:

  - Press Alt + F11 in Excel to open the VBA editor.

2. Insert a New Module:
  - Click Insert > Module to create a new module in your workbook.

3. Copy and Paste the Code:
  - Copy the VBA script code into the new module 
   
5. Run the Script:
  - Ensure the worksheet you want to run the script on is the active sheet 
  - Press F5 in the VBA editor to run the script


## Credibility

This script was developed with assistance from ChatGPT, an AI language model created by OpenAI. 
Specifically:
  - "how do I change "Set ws = ThisWorkbook.Sheets("Q1")" to run based on the current active sheet?"
    - Answer: "Set ws = ActiveSheet"
  - "Isn't there a hard limit for max calculations if you set the maxincrease and maxdecrease to arbitrary numbers?"
    - Answer: "You're correct. Setting arbitrary large initial values for maxIncrease and maxDecrease can potentially limit the calculations if the actual values exceed these initial values or if they are set incorrectly. A better approach is to use -Inf and Inf (negative infinity and positive infinity) for these variables, respectively."
    
Found Color palette code from: http://dmcritchie.mvps.org/excel/colors.htm

## NOTES & Mistakes
Had trouble figuring out why my Excel workbook wouldn't open with the macros enabled -- Realized I wasn't saving it as a macro-enabled workbook - obvious, but a mistake nonetheless

## EDITS
Reincluded the Alphabetical_Testing(finalized).xlsm file and made sure the script worked in that workbook as well -- included it into the repo in case it was needed for review
