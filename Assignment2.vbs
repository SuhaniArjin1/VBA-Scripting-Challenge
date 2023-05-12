Sub StockSummary()
Dim SumRow As Integer
Dim OpeningPrice As Double
Dim ClosingPrice As Double
Dim YearlyChange As Double
Dim PercentChange As Double
Dim TotalVolume As Double


a = Application.Worksheets.Count
    
    For j = 1 To a
        Worksheets(j).Activate
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        OpeningPrice = Range("C2").Value
        Range("I1:P1").EntireColumn.AutoFit
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row
        SumRow = 2
       

        For i = 2 To LastRow
            If Cells((i + 1), 1).Value <> Cells(i, 1).Value Then
                TotalVolume = TotalVolume + Cells(i, 7).Value
        Cells(SumRow, 9).Value = Cells(i, 1).Value
                ClosingPrice = Cells(i, 6).Value
                YearlyChange = ClosingPrice - OpeningPrice
                Cells(SumRow, 10).Value = YearlyChange
        YearlyChange = ClosingPrice - OpeningPrice
                Cells(SumRow, 10).Value = YearlyChange
                PercentChange = YearlyChange / OpeningPrice
                Cells(SumRow, 11).Value = PercentChange
                Range("K:K").NumberFormat = "0.00%"
        If Cells(SumRow, 10).Value > 0 Then
            Cells(SumRow, 10).Interior.ColorIndex = 4
        Else: Cells(SumRow, 10).Interior.ColorIndex = 3
        End If
                
        SumRow = SumRow + 1
        OpeningPrice = Cells((i + 1), 3).Value
        TotalVolume = 0
     ElseIf Cells((i + 1), 1).Value = Cells(i, 1).Value Then
                TotalVolume = TotalVolume + Cells(i, 7).Value
                Cells(SumRow, 12).Value = TotalVolume
            End If
        Next i
    Next j

End Sub

Sub MaxMinChart()
Dim GreatInc As Double
Dim GreatDec As Double
Dim GreatStock As Double


a = Application.Worksheets.Count
    
    For j = 1 To a
        Worksheets(j).Activate
        GreatInc = WorksheetFunction.Max(Range("K:K"))
        GreatDec = WorksheetFunction.Min(Range("K:K"))
        GreatStock = WorksheetFunction.Max(Range("L:L"))
        Range("O1").Value = "Ticker"
        Range("P1").Value = "Value"
        Range("N2").Value = "Greatest % Increase"
        Range("N3").Value = "Greatest % Decrease"
        Range("N4").Value = "Greatest Total Volume"
        Range("I1:P1").EntireColumn.AutoFit
        Range("P2").Value = GreatInc
        Range("P3").Value = GreatDec
        Range("P4").Value = GreatStock

        LastRow = Cells(Rows.Count, 9).End(xlUp).Row

        For i = 1 To LastRow
            If Cells(i, 11).Value = GreatInc Then
            Range("O2").Value = Cells(i, 9).Value
        End If
        If Cells(i, 11).Value = GreatDec Then
            Range("O3").Value = Cells(i, 9).Value
        End If
        If Cells(i, 12).Value = GreatStock Then
            Range("O4").Value = Cells(i, 9).Value
        End If
    Next i
Next j
End Sub








