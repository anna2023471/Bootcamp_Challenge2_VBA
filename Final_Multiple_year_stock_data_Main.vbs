Sub Multiple_year_stock_data()

Dim ws As Worksheet
Dim ticker As String
Dim Ticker_volume As Double
Dim Ticker_Yearly_Change As Double
Dim Percent_Change As Double
Dim Ticker_start_price As Double
Dim Ticker_end_price As Double

For Each ws In Worksheets

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
    
Range("J1").EntireColumn.AutoFit
Range("K1").EntireColumn.AutoFit
Range("L1").EntireColumn.AutoFit

Summary_Table_Row = 2
Last_Row = Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To Last_Row
    'Iterate through ticker and return total stock volume
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        ticker = Cells(i, 1).Value
        Ticker_volume = Ticker_volume + Cells(i, 7).Value
        Range("I" & Summary_Table_Row).Value = ticker
        Range("L" & Summary_Table_Row).Value = Ticker_volume
        Summary_Table_Row = Summary_Table_Row + 1
        Ticker_volume = 0
    Else:
        Ticker_volume = Ticker_volume + Cells(i, 7).Value
    End If
    
    'Iterate through ticker and find start price for each ticker
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value And Cells(i + 1, 1).Value <> "" Then
        Ticker_start_price = Cells(i + 1, 3).Value
    ElseIf Cells(i + 1, 1).Value = Cells(2, 1).Value Then
        Ticker_start_price = Cells(2, 3).Value
    End If
    
    'Iterate through date and find end price for each ticker
    If Cells(i + 1, 2).Value > Cells(i, 2).Value Then
        ticker = Cells(i, 1).Value
        Ticker_end_price = Cells(i + 1, 6).Value
    End If
    
    'Calculate yearly change of each ticker
    Ticker_Yearly_Change = Ticker_end_price - Ticker_start_price
    Range("J" & Summary_Table_Row).Value = Ticker_Yearly_Change
             
    'Contidionally format Yearly Change column
    If Cells(i + 1, 10).Value < 0 Then
        Cells(i + 1, 10).Interior.ColorIndex = 3
    ElseIf Cells(i + 1, 10).Value > 0 Then
        Cells(i + 1, 10).Interior.ColorIndex = 4
    ElseIf Cells(2, 10).Value < 0 Then
        Cells(2, 10).Interior.ColorIndex = 3
    ElseIf Cells(2, 10).Value > 0 Then
        Cells(2, 10).Interior.ColorIndex = 4
    End If
    
    'Calculate and format percentage change
        Percent_Change = Ticker_Yearly_Change / Ticker_start_price
        Range("K" & Summary_Table_Row).Value = Percent_Change
        Range("K" & Summary_Table_Row).Value = FormatPercent(Percent_Change, 2, vbFalse)

Next i
    
Next ws
    
End Sub