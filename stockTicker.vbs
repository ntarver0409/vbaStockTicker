Sub stockTicker()


Dim yearlyChange As Double
Dim percentChange As Double
Dim totalStockVolume As Double
Dim openPrice As Double
Dim closePrice As Double
Dim i As Long



For Each ws In Worksheets
    i = 2
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    dataTable = 2
    lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    openPrice = ws.Cells(2, 3).Value
        For j = 2 To lastRow
            If ws.Cells(j + 1, 1).Value <> ws.Cells(j, 1).Value Then
                total = total + ws.Cells(j, 7).Value
                If ws.Cells(j, 3) = 0 Then
                    For cellValue = i To j
                        If ws.Cells(cellValue, 3).Value <> 0 Then
                            i = cellValue
                            Exit For
                        End If
                    Next cellValue
                End If
                yearlyChange = ws.Cells(j, 6).Value - ws.Cells(j, 3).Value
                percentChange = (yearlyChange / ws.Cells(i, 3).Value)
                i = i + 1
                ws.Range("I" & dataTable).Value = ws.Cells(j, 1).Value
                ws.Range("J" & dataTable).Value = yearlyChange
                ws.Range("J" & dataTable).Style = "Currency"
                ws.Range("K" & dataTable).Value = percentChange
                ws.Range("K" & dataTable).Style = "Percent"
                ws.Range("L" & dataTable).Value = totalStockVolume
                With ws
                    .Columns("K:K").NumberFormat = "0.00%"
                End With
                If yearlyChange > 0 Then
                    ws.Range("J" & dataTable).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & dataTable).Interior.ColorIndex = 3
                End If
                totalStockVolume = 0
                dataTable = dataTable + 1
                yearlyChange = 0
            Else
                totalStockVolume = totalStockVolume + ws.Cells(j, 7).Value
            End If
        Next j
    Next ws
End Sub