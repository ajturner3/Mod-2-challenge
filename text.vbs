Sub stocks()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim openingPrice As Currency
    Dim closingPrice As Currency
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalStockVolume As Double
    Dim summaryPointer As Long

    For Each ws In Worksheets
        ws.Activate
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume "
    
        ws.Columns("I:L").AutoFit
    
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        totalStockVolume = 0
        openingPrice = ws.Cells(2, "C").Value
        summaryPointer = 2
    
    For i = 2 To lastRow
        If ws.Cells(i, "A").Value = ws.Cells(i + 1, "A").Value Then
           totalStockVolume = totalStockVolume + ws.Cells(i, "G").Value
            Else
                closingPrice = ws.Cells(i, "F").Value
                yearlyChange = closingPrice - openingPrice
                totalStockVolume = totalStockVolume + ws.Cells(i, "G").Value
                percentChange = (yearlyChange / openingPrice) * 100
                ws.Cells(summaryPointer, "I").Value = ws.Cells(i, "A").Value
                ws.Cells(summaryPointer, "J").Value = yearlyChange
                ws.Cells(summaryPointer, "K").Value = percentChange
                ws.Cells(summaryPointer, "L").Value = totalStockVolume
        
                If yearlyChange > 0 Then
                    ws.Range("J" & summaryPointer).Interior.Color = vbGreen
                ElseIf yearlyChange < 0 Then
                    ws.Range("J" & summaryPointer).Interior.Color = vbRed
                Else
                    ws.Range("J" & summaryPointer).Interior.Color = vbWhite
                End If
        
                totalStockVolume = 0
                openingPrice = ws.Cells(i + 1, "C").Value
                summaryPointer = summaryPointer + 1
        
            End If
        Next i
    Next ws
    MsgBox ("complete")
End Sub


