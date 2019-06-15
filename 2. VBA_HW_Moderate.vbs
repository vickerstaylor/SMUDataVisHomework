Sub TickerMath_Moderate()

    Dim ws As Worksheet

    For Each ws In Worksheets
    
        Dim ticker_name As String

        Dim ticker_total As Double
        ticker_total = 0

        Dim open_price As Double
        open_price = 0

        Dim close_price As Double
        close_price = 0

        Dim price_diff As Double
        price_diff = 0

        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
        ws.Range("I1").Columns.AutoFit
        ws.Range("J1").Columns.AutoFit
        ws.Range("K1").Columns.AutoFit
        ws.Range("L1").Columns.AutoFit
        
        ws.Cells(1, 9) = "Ticker"
        ws.Cells(1, 10) = "Yearly Change"
        ws.Cells(1, 11) = "Percent Change"
        ws.Cells(1, 12) = "Total Stock Volume"

        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        For i = 2 To lastrow

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker_name = ws.Cells(i, 1).Value
                ticker_total = ticker_total + ws.Cells(i, 7).Value
                ws.Range("I" & Summary_Table_Row).Value = ticker_name
                ws.Range("J" & Summary_Table_Row).Value = price_diff
                ws.Range("K" & Summary_Table_Row).Value = pct_change
                ws.Range("L" & Summary_Table_Row).Value = ticker_total
                Summary_Table_Row = Summary_Table_Row + 1
                close_price = ws.Cells(i + 1, 6).Value
                open_price = ws.Cells(i, 3).Value
                ticker_total = 0
            Else
                ticker_total = ticker_total + ws.Cells(i, 7).Value
                open_price = ws.Cells(i + 1, 3).Value
                close_price = ws.Cells(i + 1, 6).Value
                price_diff = close_price - open_price
                pct_change = ((close_price - open_price) / open_price) * 100
    
            End If
                
        Next i
    
    Next ws

End Sub

