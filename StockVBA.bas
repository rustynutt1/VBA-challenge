Attribute VB_Name = "Stock"
Sub StockData()

    Dim ws As Worksheet
    Dim LastRow As Long
    Dim Ticker As String
    Dim MinPrice As Double
    Dim ClosingPrice As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalVolume As Double
    Dim SummaryRow As Integer
    
    ' Loop through all worksheets
    For Each ws In ThisWorkbook.Worksheets
    
        ' Initialize variables
        SummaryRow = 2
        TotalVolume = 0
        
        ' Find the last row of data in the current worksheet
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Loop through rows and perform calculations
        For i = 2 To LastRow
        
            ' Check if the current row's ticker symbol is different from the previous row
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                ' Set min price for the current ticker symbol
                MinPrice = ws.Cells(i, 3).Value
            End If
            
            ' Add up total stock volume
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value
            
            ' Check if the current row's ticker symbol is different from the next row
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                ' Set closing price for the current ticker symbol
                ClosingPrice = ws.Cells(i, 6).Value
                
                ' Calculate yearly change and percent change
                YearlyChange = ClosingPrice - MinPrice
                If MinPrice <> 0 Then
                    PercentChange = (YearlyChange / MinPrice)
                Else
                    PercentChange = 0
                End If
                
                ' Output data to summary table
                ws.Cells(SummaryRow, 9).Value = ws.Cells(i, 1).Value ' Ticker Symbol
                ws.Cells(SummaryRow, 10).Value = YearlyChange ' Yearly Change
                ws.Cells(SummaryRow, 11).Value = Format(PercentChange, "Percent") ' Percent Change
                ws.Cells(SummaryRow, 12).Value = TotalVolume ' Total Stock Volume
                
                ' Color code positive and negative changes
                If YearlyChange >= 0 Then
                    ws.Cells(SummaryRow, 10).Interior.ColorIndex = 4 ' Green for positive change
                Else
                    ws.Cells(SummaryRow, 10).Interior.ColorIndex = 3 ' Red for negative change
                End If
                
                ' Move to the next row in the summary table
                SummaryRow = SummaryRow + 1
                
                ' Reset total volume for the next ticker symbol
                TotalVolume = 0
            End If
        
        Next i
        
        ' Find the last row of data in the summary table
        LastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        ' Find the stock with the greatest percent increase, decrease, and total volume
        Dim MaxIncrease As Double, MaxDecrease As Double, MaxVolume As Double
        Dim MaxIncreaseTicker As String, MaxDecreaseTicker As String, MaxVolumeTicker As String
        
        MaxIncrease = WorksheetFunction.Max(ws.Range("K2:K" & LastRow))
        MaxDecrease = WorksheetFunction.Min(ws.Range("K2:K" & LastRow))
        MaxVolume = WorksheetFunction.Max(ws.Range("L2:L" & LastRow))
        
        MaxIncreaseTicker = ws.Cells(Application.Match(MaxIncrease, ws.Range("K2:K" & LastRow), 0) + 1, 9).Value
        MaxDecreaseTicker = ws.Cells(Application.Match(MaxDecrease, ws.Range("K2:K" & LastRow), 0) + 1, 9).Value
        MaxVolumeTicker = ws.Cells(Application.Match(MaxVolume, ws.Range("L2:L" & LastRow), 0) + 1, 9).Value
        
        ' Output results for greatest percent increase, decrease, and total volume
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        
        ws.Cells(2, 15).Value = MaxIncreaseTicker
        ws.Cells(3, 15).Value = MaxDecreaseTicker
        ws.Cells(4, 15).Value = MaxVolumeTicker
        
        ws.Cells(2, 16).Value = MaxIncrease & "%"
        ws.Cells(3, 16).Value = MaxDecrease & "%"
        ws.Cells(4, 16).Value = MaxVolume
        
        ' Apply conditional formatting for positive and negative changes
        ws.Range("K2:K" & LastRow).FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="0"
        ws.Range("K2:K" & LastRow).FormatConditions(1).Interior.ColorIndex = 4 ' Green for positive change
        
        ws.Range("K2:K" & LastRow).FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
        ws.Range("K2:K" & LastRow).FormatConditions(2).Interior.ColorIndex = 3 ' Red for negative change
        
    Next ws
    
End Sub
