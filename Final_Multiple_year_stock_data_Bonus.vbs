Sub Multiple_year_stock_data()

Dim ticker As String
Dim PercentIncrease As Double
Dim PercentageDecrease As Double
Dim TotalVolume As Double

Increase_Table_Row = 2
Decrease_Table_Row = 3
Volume_Table_Row = 4

Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greates Total Volume"

Range("O1").EntireColumn.AutoFit
Last_Row = Cells(Rows.Count, 1).End(xlUp).Row

'Return max and min values
PercentIncrease = Application.WorksheetFunction.Max(Range("K2:K4000"))
Range("Q" & Increase_Table_Row).Value = PercentIncrease
PercentDecrease = Application.WorksheetFunction.Min(Range("K2:K4000"))
Range("Q" & Decrease_Table_Row).Value = PercentDecrease
TotalVolume = Application.WorksheetFunction.Max(Range("L2:L4000"))
Range("Q" & Volume_Table_Row).Value = TotalVolume

'Format percentages
Range("Q" & Increase_Table_Row).Value = FormatPercent(PercentIncrease, 2)
Range("Q" & Decrease_Table_Row).Value = FormatPercent(PercentDecrease, 2)

'Return corresponding ticker
For i = 1 To Last_Row
    If Cells(i + 1, 11).Value = PercentIncrease Then
        ticker = Cells(i + 1, 9).Value
        Range("P" & Increase_Table_Row).Value = ticker
    ElseIf Cells(i + 1, 11).Value = PercentDecrease Then
        ticker = Cells(i + 1, 9).Value
        Range("P" & Decrease_Table_Row).Value = ticker
    ElseIf Cells(i + 1, 12).Value = TotalVolume Then
        ticker = Cells(i + 1, 9).Value
        Range("P" & Volume_Table_Row).Value = ticker
    End If
Next i

End Sub
