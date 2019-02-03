Sub alphabetical_testing()
'Create a script that will loop through each year of stock data and grab the total amount of volume each stock had over the year.
'You will also need to display the ticker symbol to coincide with the total volume.
'Identifying Variables:
   Dim Ticker As String
   Dim Total_Stock_Volume As LongLong
   Dim Summary_Table_Row As Integer
   'Total_Stock_Volume:
   Total_Stock_Volume = 0
   'Insert Headers for Column 9 & 10:
       For Each ws In Worksheets
       ws.Cells(1, 9).Value = "Ticker"
       ws.Cells(1, 10).Value = "Total Stock Volume"
       'Knowing the Last Row:
       LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
       Summary_Table_Row = 2
       'Looping through Stocks:
           For i = 2 To LastRow
           'Conditionals:
               If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
               'Ideitify Tickers
               Ticker = ws.Cells(i, 1).Value
               'Calculation of Total Stock Vol.
               Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
               'Input result in Summary Table
               ws.Range("I" & Summary_Table_Row).Value = Ticker
               ws.Range("J" & Summary_Table_Row).Value = Total_Stock_Volume
               'Add one to the Summary Table:
               Summary_Table_Row = Summary_Table_Row + 1
               'Reset for ticker:
               Total_Stock_Volume = 0
               Else
               Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
End If
Next i
Next ws
End Sub
