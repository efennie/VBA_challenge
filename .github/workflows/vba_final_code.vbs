Sub stock_market_review()

' create variables
Dim LastRow As Long
Dim ticker As String
Dim open_price As Double
Dim close_price As Double
Dim stock_volume As LongLong
Dim summary_table_row As Long
Dim year_change As Double

summary_table_row = 2

' figure out how many rows there are per sheet
LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
'sanity check w/ message box
'MsgBox (LastRow)

' set up summary table headings

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

' change percent change column formatting into percentage
Range("K2:K" & LastRow).NumberFormat = "0.00%"

' Adjust column width
Range("A:P").Columns("A:P").AutoFit


' set up opening prices for first ticker

open_price = Cells(2, 3).Value

' run the first row's info:

For i = 2 To LastRow

    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
        ticker = Cells(i, 1).Value
        
        ' input ticker into summary table
        Range("I" & summary_table_row).Value = ticker
        
        ' input closing prices
        close_price = Cells(i, 6).Value
        
        year_change = close_price - open_price
        
        Range("J" & summary_table_row).Value = year_change
        
        ' add to the stock_volume
        stock_volume = stock_volume + Cells(i, 7).Value
        
        ' input stock total into summary table
        Range("L" & summary_table_row).Value = stock_volume
        
        ' calculate percent change and input it into the table
        year_change = ((close_price - open_price) / (open_price))
        Range("K" & summary_table_row).Value = year_change
        
         'conditional formatting
  
            If year_change >= 0 Then
        
                Range("J" & summary_table_row).Interior.ColorIndex = 4
                
            Else
                Range("J" & summary_table_row).Interior.ColorIndex = 3
                
            End If
        
        summary_table_row = summary_table_row + 1
        stock_volume = 0
        
' store the new open_price as a variable

        open_price = Cells(i + 1, 3).Value
    ' take care of open prices that are 0
        
            If open_price = 0 Then
    
                For j = (i + 1) To LastRow
                    
                    If Cells(j, 3) <> 0 Then
                        open_price = Cells(j, 3)
                        
                    End If
                    
                    Next j
                    
            End If
            
    ' If cell right before it is the same:
    Else
        
'   add stock_volume to total
        stock_volume = stock_volume + Cells(i, 7).Value
        
      
    End If
    
    Next i

MsgBox ("Done")

End Sub
