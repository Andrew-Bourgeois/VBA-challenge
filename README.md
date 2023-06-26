# stock-analysis
Creating new repository called "stock-analysis" per the UNC-CH-DA Week 2-Module 2.2 instructions.

This the UNC Chapel Hill Data Analytics Modlue 2 Challenge repository.

Background: Use VBA scripting to analyze generated stock market data.

Instructions: Create a script which loops through the provided Excel workbook of annual stock data. Script should iterate through each stock and output the following in a new set of data tables (either on the same worksheet or a new worksheet):
  *  The ticker symbol
  *  Annual change between stock price of stock on the closing day of the year and the opening day of the same year.
      - NOTE: positive changes by highlighting the cell in green, negative changes by highlighting the cell in red.
  *  Percentage change of stock price from opening to closing day of the year
      - NOTE: added additional conditional formatting so negative percentages will appear RED 
  *  Total volume of stock from opening day to the closing day of the year.
  *  Add additional scripting to track and report in an additional data table, the following:
      -  Greatest % increase
      -  Greatest % decrease
      -  Greates total volume
  *  Ensure the script will iterate and run over each yearly worksheet in the workbook with a single call.

Citations:
  * For looping through worksheets, referenced the folling site: https://support.microsoft.com/en-us/topic/macro-to-loop-through-all-worksheets-in-a-workbook-feef14e3-97cf-00e2-538b-5da40186e2b0
  * For Autofitting Columns using VBA, referenced the following site: https://learn.microsoft.com/en-us/office/vba/api/excel.range.autofit
  * For converting Percent changes column to percentage format: https://www.educba.com/vba-number-format/
