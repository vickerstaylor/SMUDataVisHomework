Sub TickerMath_Easy()

    Dim ws As Worksheet

    For Each ws In Worksheets
    
        Dim ticker_name As String

        Dim ticker_total As Double
        ticker_total = 0

        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
        ws.Range("I1").Columns.AutoFit
        ws.Range("J1").Columns.AutoFit
        
        ws.Cells(1, 9) = "Ticker"
        ws.Cells(1, 10) = "Total Stock Volume"

        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        For i = 2 To lastrow

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker_name = ws.Cells(i, 1).Value
                ticker_total = ticker_total + ws.Cells(i, 7).Value
                ws.Range("I" & Summary_Table_Row).Value = ticker_name
                ws.Range("J" & Summary_Table_Row).Value = ticker_total
                Summary_Table_Row = Summary_Table_Row + 1
                ticker_total = 0
            Else
                ticker_total = ticker_total + ws.Cells(i, 7).Value
    
            End If
                
        Next i
    
    Next ws

End Sub
