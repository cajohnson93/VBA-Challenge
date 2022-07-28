Sub Module2_Challenge()

For Each ws In Worksheets

'set headers for columns

    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"

'set variables

    Dim RowCount As Long
    Dim Ticker As String
    Dim YrChange As Double
    Dim PerChange As Double
    Dim Vol As Double
    Dim StockOp As Double
    Dim StockCl As Double
    Dim GrtIncr As Double
    Dim GrtDecr As Double
    Dim GrtVol As Double
    Dim Summary As Long
    
RowCount = ws.Cells(Rows.Count, "1").End(xlUp).Row
Vol = 0
Summary = 2


'set rows for loops

For i = 2 To RowCount

'set ticker name

                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
'set ticker
                Ticker = ws.Cells(i, 1).Value
                Vol = Vol + ws.Cells(i, 7).Value
                 
'Calculate Yearly Change
            

            ws.Range("I" & Summary).Value = Ticker
            ws.Range("L" & Summary).Value = Volume

        Vol = 0

        StockCl = ws.Cells(i, 6)

 
       
        If StockOpen = 0 Then
            YrChange = 0
            PerChange = 0
        Else:
            YrChange = StockCl - StockOp
            PerChange = (StockCl - StockOp) / StockOp
        End If

    
            ws.Range("J" & Summary).Value = YrChange
            ws.Range("K" & Summary).Value = PerChange
            ws.Range("K" & Summary).Style = "Percent"
            ws.Range("K" & Summary).NumberFormat = "0.00%"

            Summary = Summary + 1

    ElseIf ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1) Then
         StockOp = ws.Cells(i, 3)


    Else: Vol = Vol + ws.Cells(i, 7).Value

    End If
    
    
  Next i
    
    For i = 2 To RowCount
    

'set colors

            If ws.Cells("1, 10" & i).Value < 0 Then
            ws.Cells("1, 10" & i).Interior.ColorIndex = 3
            
            ElseIf ws.Cells("1, 10" & i).Value > 0 Then
            ws.Cells("1, 10" & i).Interior.ColorIndex = 4

                    
    End If

Next i

ws.Columns("A:Q").AutoFit

Next ws

End Sub

