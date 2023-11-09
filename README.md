# VBA-challenge
VBA scripting to analyze generated stock market data.
Sub CalculateStockMetricsForAllWorksheets()
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim Ticker As String
    Dim YearOpen As Double
    Dim YearClose As Double
    Dim YearChange As Double
    Dim PercentChange As Double
    Dim TotalVolume As Double
    Dim OutputRow As Long
    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim GreatestVolume As Double
    Dim GreatestIncreaseTicker As String
    Dim GreatestDecreaseTicker As String
    Dim GreatestVolumeTicker As String
    
    ' Loop through all worksheets in the workbook
    For Each ws In Worksheets
        ' Initialize variables for each worksheet
        YearOpen = 0
        YearClose = 0
        TotalVolume = 0
        OutputRow = 2  ' Start output from row 2
        GreatestIncrease = 0
        GreatestDecrease = 0
        GreatestVolume = 0
        
        ' Find the last row with data in the current worksheet
        LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' Loop through the data in the current worksheet
        For i = 2 To LastRow
            If Ticker <> ws.Cells(i, 1).Value Then
                ' Output results for the previous stock
                If Ticker <> "" Then
                    YearChange = YearClose - YearOpen
                    If YearOpen <> 0 Then
                        PercentChange = (YearChange / YearOpen) * 100
                    Else
                        PercentChange = 0
                    End If
                    ws.Cells(OutputRow, 10).Value = Ticker
                    ws.Cells(OutputRow, 11).Value = YearOpen
                    ws.Cells(OutputRow, 12).Value = YearClose
                    ws.Cells(OutputRow, 13).Value = YearChange
                    ws.Cells(OutputRow, 14).Value = PercentChange
                    ws.Cells(OutputRow, 15).Value = TotalVolume
                    
                    ' Apply conditional formatting for yearly change
                    If YearChange < 0 Then
                        ws.Cells(OutputRow, 13).Interior.Color = RGB(255, 0, 0) ' Red for negative change
                    Else
                        ws.Cells(OutputRow, 13).Interior.Color = RGB(0, 255, 0) ' Green for positive change
                    End If
                    
                    ' Update the greatest increase, decrease, and volume
                    If PercentChange > GreatestIncrease Then
                        GreatestIncrease = PercentChange
                        GreatestIncreaseTicker = Ticker
                    ElseIf PercentChange < GreatestDecrease Then
                        GreatestDecrease = PercentChange
                        GreatestDecreaseTicker = Ticker
                    End If
                    
                    If TotalVolume > GreatestVolume Then
                        GreatestVolume = TotalVolume
                        GreatestVolumeTicker = Ticker
                    End If
                    
                    OutputRow = OutputRow + 1 ' Move to the next row
                End If
                
                ' Update the current stock information
                Ticker = ws.Cells(i, 1).Value
                YearOpen = ws.Cells(i, 3).Value
                TotalVolume = 0
            End If
            
            ' Accumulate total volume
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value
            
            ' Update the closing price at the end of the year
            YearClose = ws.Cells(i, 6).Value
        Next i
        
        ' Output results for the last stock in the current worksheet
        YearChange = YearClose - YearOpen
        If YearOpen <> 0 Then
            PercentChange = (YearChange / YearOpen) * 100
        Else
            PercentChange = 0
        End If
        ws.Cells(OutputRow, 10).Value = Ticker
        ws.Cells(OutputRow, 11).Value = YearOpen
        ws.Cells(OutputRow, 12).Value = YearClose
        ws.Cells(OutputRow, 13).Value = YearChange
        ws.Cells(OutputRow, 14).Value = PercentChange
        ws.Cells(OutputRow, 15).Value = TotalVolume
        
        ' Apply conditional formatting for yearly change in the last row
        If YearChange < 0 Then
            ws.Cells(OutputRow, 13).Interior.Color = RGB(255, 0, 0) ' Red for negative change
        Else
            ws.Cells(OutputRow, 13).Interior.Color = RGB(0, 255, 0) ' Green for positive change
        End If
        
        ' Display the greatest increase, decrease, and volume in columns Q1:Q5 for the current worksheet
        ws.Cells(2, 17).Value = "Greatest % Increase"
        ws.Cells(3, 17).Value = GreatestIncrease
        ws.Cells(4, 17).Value = GreatestIncreaseTicker
        ws.Cells(5, 17).Value = "Greatest % Decrease"
        ws.Cells(6, 17).Value = GreatestDecrease
        ws.Cells(7, 17).Value = GreatestDecreaseTicker
        ws.Cells(8, 17).Value = "Greatest Total Volume"
        ws.Cells(9, 17).Value = GreatestVolume
        ws.Cells(10, 17).Value = GreatestVolumeTicker
    Next ws
End Sub
