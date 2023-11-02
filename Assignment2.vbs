Sub StockData()

    Dim LastRow As Long
    Dim i As Long
    Dim StartRow As Long
    Dim TickerSymbol As String
    Dim YearlyChange, PercentageChange, TotalVolume, OpenPrice, ClosePrice As Double
    Dim GreatestIncrease, GreatestDecrease, GreatestTotalVolume As Double
    Dim GreatestIncreaseTicker, GreatestDecreaseTicker, GreatestVolumeTicker As String

    ' Find the last row
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row

    ' Set the initial row for summary output
    StartRow = 2
    
    ' Set headers
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percentage Change"
    Cells(1, 12).Value = "Total Stock Volume"

    ' Loop through stocks
    For i = 2 To LastRow
        
        ' Check if we are at the start of a new stock ticker
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            TickerSymbol = Cells(i, 1).Value
            ClosePrice = Cells(i, 6).Value
            YearlyChange = ClosePrice - OpenPrice
            If OpenPrice <> 0 Then
                PercentageChange = (YearlyChange / OpenPrice) * 100
            Else
                PercentageChange = 0
            End If
            'TotalVolume = TotalVolume + Cells(i, 7).Value
            

            ' Output the data
            Cells(StartRow, 9).Value = TickerSymbol
            Cells(StartRow, 10).Value = YearlyChange
            Cells(StartRow, 11).Value = PercentageChange & "%"
            Cells(StartRow, 12).Value = TotalVolume
            
            ' Conditional Formatting
            If YearlyChange > 0 Then
                Cells(StartRow, 10).Interior.Color = vbGreen
            ElseIf YearlyChange < 0 Then
                Cells(StartRow, 10).Interior.Color = vbRed
            End If

            ' Check for greatest metrics
            If YearlyChange > GreatestIncrease Or StartRow = 2 Then
                GreatestIncrease = YearlyChange
                GreatestIncreaseTicker = TickerSymbol
            End If

            If YearlyChange < GreatestDecrease Or StartRow = 2 Then
                GreatestDecrease = YearlyChange
                GreatestDecreaseTicker = TickerSymbol
            End If

            If TotalVolume > GreatestTotalVolume Or StartRow = 2 Then
            GreatestTotalVolume = TotalVolume
            GreatestVolumeTicker = TickerSymbol
            End If

            ' Reset values for the next stock
            TotalVolume = 0
            StartRow = StartRow + 1
        Else
            ' Not at the end of the current stock ticker, accumulate the volume
            If i = 2 Then
                OpenPrice = Cells(i, 3).Value
            End If
            'TotalVolume = TotalVolume + Cells(i, 7).Value
        End If
    Next i

    ' Output the greatest metrics
    Cells(2, 14).Value = "Greatest % Increase"
    Cells(3, 14).Value = "Greatest % Decrease"
    Cells(4, 14).Value = "Greatest Total Volume"
    Cells(2, 15).Value = GreatestIncreaseTicker
    Cells(3, 15).Value = GreatestDecreaseTicker
    Cells(4, 15).Value = GreatestVolumeTicker
    Cells(2, 16).Value = GreatestIncrease
    Cells(3, 16).Value = GreatestDecrease
    Cells(4, 16).Value = GreatestTotalVolume

End Sub

