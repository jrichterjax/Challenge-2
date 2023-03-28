# VBA-challenge

This VBA script summarizes three years' stock data according to the change in opening and closing values, percent change, and total stock volume, as well as highlights the stock with the greatest percent change, least percent change, and greatest total stock volume per year.

The code first assigns headers and summary table labels. 

Variables are established to hold yearly change, percent change, total stock volume, summary row, and last row. The variable "yearlyCounter" is used to determine the starting value for each stock and save it as the variable "openValue" for later use.

A For loop is used to compare each row's ticker with the next row's ticker. If they are equal, total stock volume is compounded and one is added to the yearlyCounter variable. When yearly counter equals one, the open value is stored. 

If the tickers are not equal, the summary columns are formatted and populated with the ticker, the calculated yearly change, percent change, and total stock volume, and variables are reset or updated as needed. 

To find the greatest percent increase/decrease and greatest total volume, the first ticker's percent change (column k) and total volume (column l) are assigned to variables. These variables are then compared to the next ticker's values, and depending on whether the next values are greater than or less than, the variable remains the same or is reassigned.

After all ticker values have been compared, the remaining values for each variable are assigned to the appropriate field in the summary table.

The script then autofits the columns and completes these actions on all worksheets. 

When completed, a message box stating "Summaries Completed" appears.