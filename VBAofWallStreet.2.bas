Attribute VB_Name = "Module2"
Sub tickerConditional()

For Each ws In Worksheets

Dim yearlyChange As Double
Dim totalVolume As Double
Dim percentChange As Double
Dim summaryTableRow As Double
Dim tickerSymbol As String
Dim openingValue As Long
Dim closingValue As Long
lastA = ws.Cells(Rows.Count, "A").End(xlUp).Row
lastJ = ws.Cells(Rows.Count, "J").End(xlUp).Row
lastM = ws.Cells(Rows.Count, "M").End(xlUp).Row
ws.Range("j1").Value = "Ticker Symbol"
ws.Range("k1").Value = "Annual Closing Value"
ws.Range("l1").Value = "Annual Opening Value"
ws.Range("m1").Value = "Yearly Change"
ws.Range("n1").Value = "Percent Change"
ws.Range("o1").Value = "Total Annual Volume"
totalVolume = 0
summaryTableRow = 2

For i = 2 To lastA
    totalVolume = totalVolume + ws.Cells(i, "g").Value
    ws.Range("o" & summaryTableRow).Value = totalVolume
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        tickerSymbol = ws.Cells(i, "a").Value
        closingValue = ws.Cells(i, "f").Value
        ws.Range("j" & summaryTableRow).Value = tickerSymbol
        ws.Range("k" & summaryTableRow).Value = closingValue
        summaryTableRow = summaryTableRow + 1
        totalVolume = 0
    End If
Next i
totalVolume = 0
summaryTableRow = 2

For j = 2 To lastA
    If ws.Cells(j - 1, "a").Value <> ws.Cells(j, "a").Value Then
        openingValue = ws.Cells(j, "c").Value
        ws.Range("l" & summaryTableRow).Value = openingValue
        summaryTableRow = summaryTableRow + 1
    End If
Next j
summaryTableRow = 2

For k = 2 To lastJ
    yearlyChange = ws.Cells(k, "k").Value - ws.Cells(k, "l").Value
    ws.Range("m" & summaryTableRow).Value = yearlyChange
    If ws.Cells(k, "l").Value = 0 Then
        percentChange = 0
    Else
        percentChange = (ws.Cells(k, "m").Value / ws.Cells(k, "l").Value)
    End If
    ws.Range("n" & summaryTableRow).Value = percentChange
    Dim percentChangeRange As Range
    Set percentChangeRange = ws.Range("n" & summaryTableRow)
    percentChangeRange.NumberFormat = "0.00%"
    summaryTableRow = summaryTableRow + 1
Next k

Dim yearlyChangeRange As Range
Set yearlyChangeRange = ws.Range("m2", lastJ)
Dim condition1 As FormatCondition, condition2 As FormatCondition, condition3 As FormatCondition
Set condition1 = ws.yearlyChangeRange.FormatConditions.Add(xlCellValue, xlGreater, "0")
Set condition2 = ws.yearlyChangeRange.FormatConditions.Add(xlCellValue, xlLess, "0")

With condition1
    .Interior.Color = vbGreen
End With
With condition2
    .Interior.Color = vb = vbRed
End With

Next ws
End Sub

