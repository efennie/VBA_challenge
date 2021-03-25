# VBA_challenge

#This is a script used to loop through stock market data to calculate the stock price changes within a year, calculate the total stock volume, and the percent change from opening stock prices to closing stock prices.

' figure out how many rows there are per sheet
LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
' in class we used this row finder but with ".ws" ahead of it bc it was per sheet
'sanity check w/ message box
MsgBox (LastRow)

' set up summary table headings and create a counter
'   to track which row we are on
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

' Adjust column width
Worksheets("A").Columns("A:L").AutoFit
