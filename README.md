# VBA-Module-2-Challenge

This is an insight as to what I did to achieve the required outcomes for this Module Assignment.

The key step is creating a line that determines the last row of a sheet and grab all of the requred data for every sheet in the workbook.

The next key step is to combine all of the values under the same ticker for a given year by creating a loop. If a ticker name remains the same then add the volume traded for a given day into the total

Also determine the opening price for a year and the closing price and determine the net change in terms of price and percentage change. 

Also format to which the net loss are highlisted in a red box and net profit are highlighted in green.

The percentage values are also formatted in manner where it has 2 decimal places and a % sign at the end.

For everytime the ticker name remains the same, the values are all added up into the total stock volume for the tickers mentioned. Then using multiple loops, the tickers for the greatest % gain and loss and the greatest total volume are collected. These are all initially applied on a ws level however the ws scrip will repeat for every sheet within the workbook. 