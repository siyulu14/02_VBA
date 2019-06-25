Attribute VB_Name = "Module1"
' Moderate
' Create a script that will loop through all the stocks and take the following info.
' Yearly change from what the stock opened the year at to what the closing price was.
' The percent change from the what it opened the year at to what it closed.
' The total Volume of the stock
' Ticker symbol

Sub Ticker_counter():
      'Loop through all sheets
      
      For Each ws In Worksheets
            
            ' define variables
            Dim op_price, cl_price, delta, delta_pct, total_vol As Double
            Dim this_ticker, next_ticker As String
            op_price = Cells(2, "C").Value
            cl_price = 0
            delta = 0
            delta_pct = 0
            
            ' Tracks the location for each ticker in the summary table
            Dim summary_row As Integer
            summary_row = 2
            
            LastRow = Cells(Rows.Count, 1).End(xlUp).Row
            ' loop through all tickers
            For i = 2 To LastRow
                  
                this_ticker = Cells(i, "A").Value
                next_ticker = Cells(i + 1, "A").Value
                total_vol = total_vol + Cells(i, "G").Value
                
                ' check if ticker symbol changed
                If this_ticker <> next_ticker Then
                      
                      ' get close price
                      cl_price = Cells(i, "F").Value
                      
                      ' get yearly change, yearly change percent
                      delta = cl_price - op_price
                      If delta = 0 Then
                             delta_pct = 0
                      Else
                             delta_pct = (CDbl(delta) / CDbl(op_price))
                      End If
                      
                       ' Write values to the Summary Table (whatever your columns are):
                      Cells(summary_row, "I").Value = this_ticker
                      Cells(summary_row, "J").Value = delta
                      Cells(summary_row, "K").Value = delta_pct
                      Cells(summary_row, "L").Value = total_vol
                      Cells(summary_row, "K").NumberFormat = "00.00%"
                       ' Reset the variables for the next ticker symbol
                       op_price = Cells(i + 1, "C").Value
                       total_vol = 0
                       summary_row = summary_row + 1
                       
                End If
           Next i
        
         ' Write the header to the colums
         Cells(1, "I") = "Ticker"
         Cells(1, "J") = "Yearly Change"
         Cells(1, "K") = "Pecent Change"
         Cells(1, "L") = "Total Stock Volume"
   
        ' Color the percent
        LastRow = Cells(Rows.Count, 9).End(xlUp).Row
        For i = 2 To LastRow
            If Cells(i, "J") < 0 Then
                Cells(i, "J").Interior.ColorIndex = 3
            Else
                Cells(i, "J").Interior.ColorIndex = 4
       
            End If
        Next i
        
        ws.Activate
        
    Next ws
    
End Sub

