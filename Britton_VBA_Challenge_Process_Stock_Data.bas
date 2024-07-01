Attribute VB_Name = "Module1"
Sub ProcessStockData()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openPrice As Double, closePrice As Double, totalVolume As Double
    Dim change As Double, percentChange As Double
    Dim i As Long, outputRow As Long
    Dim tickerStartRow As Long
    Dim formatRange As Range
    
    ' Variables for greatest values
    Dim greatestIncrease As Double, greatestDecrease As Double, greatestVolume As Double
    Dim tickerIncrease As String, tickerDecrease As String, tickerVolume As String
    
    ' Loop through each worksheet
    For Each ws In ThisWorkbook.Worksheets
        ' Find the last row with data in column A
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' Write headers in columns I, J, K, L
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        outputRow = 2
        tickerStartRow = 2
        
        ' Initialize first ticker and total volume
        ' totalVolume needs to be reset to zero here to calculate the next worksheet (quarter) correctly
        ticker = ws.Cells(2, 1).Value
        totalVolume = 0
        
        ' Initialize greatest values
        ' Set these variables to 0 so you can compare percentChange and totalVolume
        ' and reassign with the correct value for that iteration
        greatestIncrease = 0
        greatestDecrease = 0
        greatestVolume = 0
        
        ' Loop through each row, starting from row 2 (skipping header)
        For i = 2 To lastRow + 1
            ' Check if we're still within the same ticker, or if we've reached a new ticker (or end of data)
            If i > lastRow Or ws.Cells(i, 1).Value <> ticker Then
                
                openPrice = ws.Cells(tickerStartRow, 3).Value
                closePrice = ws.Cells(i - 1, 6).Value
                change = closePrice - openPrice
                
                If openPrice <> 0 Then
                    percentChange = change / openPrice
                Else
                    percentChange = 0
                End If
                
                ' Write the results in the appropriate columns identified above
                ws.Cells(outputRow, 9).Value = ticker
                ws.Cells(outputRow, 10).Value = change
                ws.Cells(outputRow, 11).Value = percentChange
                ws.Cells(outputRow, 12).Value = totalVolume
                
                ' Format percent change as percentage
                'ws.Cells(outputRow, 11).NumberFormat = "0.00%"
                
                ' Check for greatest values
                If percentChange > greatestIncrease Then
                    greatestIncrease = percentChange
                    tickerIncrease = ticker
                End If
                
                If percentChange < greatestDecrease Or greatestDecrease = 0 Then
                    greatestDecrease = percentChange
                    tickerDecrease = ticker
                End If
                
                If totalVolume > greatestVolume Then
                    greatestVolume = totalVolume
                    tickerVolume = ticker
                End If
                
                ' Reset variables for next ticker
                outputRow = outputRow + 1
                
                If i <= lastRow Then
                    ticker = ws.Cells(i, 1).Value
                    tickerStartRow = i
                    totalVolume = 0
                End If
            End If
            
            ' Accumulate the volume
            If i <= lastRow Then
                totalVolume = totalVolume + ws.Cells(i, 7).Value
            End If
        Next i
        
    Dim lastOutputRow As Long
    Dim formatQuarterChange As Range
    Dim formatPercentChange As Range
    
        lastOutputRow = outputRow - 1
        
        ' Apply conditional formatting to Quarterly Change Column
        Set formatQuarterChange = ws.Range("J2:J" & lastOutputRow)
        formatQuarterChange.FormatConditions.Delete
        
        ' Format column to two decimal places
        formatQuarterChange.NumberFormat = "0.00"
        
        ' Format positive changes to green
        formatQuarterChange.FormatConditions.Add(xlCellValue, xlGreater, "= 0").Interior.Color = RGB(0, 255, 0)
        
        ' Format negative changes to red
        formatQuarterChange.FormatConditions.Add(xlCellValue, xlLess, "= 0").Interior.Color = RGB(255, 0, 0)
        
        ' Apply conditional formatting to Percent Change Column
        ' For the screeneshot to look exactly like the example on the Module Homework page I would need to
        ' comment out or remove the Format... lines below. Except, I would have to leave the set range and delete
        ' lines for a spreadsheet that had previously ran the script with formatPercentChange active.
        
        Set formatPercentChange = ws.Range("K2:K" & lastOutputRow)
        formatPercentChange.FormatConditions.Delete

        ' Format column to Percent
        formatPercentChange.NumberFormat = "0.00%"

        ' Format positive changes to green
        formatPercentChange.FormatConditions.Add(xlCellValue, xlGreater, "= 0").Interior.Color = RGB(0, 255, 0)

        ' Format negative changes to red
        formatPercentChange.FormatConditions.Add(xlCellValue, xlLess, "= 0").Interior.Color = RGB(255, 0, 0)
        
        ' Write greatest values
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        
        ws.Cells(2, 15).Value = tickerIncrease
        ws.Cells(2, 16).Value = greatestIncrease
        ws.Cells(2, 16).NumberFormat = "0.00%"
        
        ws.Cells(3, 15).Value = tickerDecrease
        ws.Cells(3, 16).Value = greatestDecrease
        ws.Cells(3, 16).NumberFormat = "0.00%"
        
        ws.Cells(4, 15).Value = tickerVolume
        ws.Cells(4, 16).Value = greatestVolume
        ws.Cells(4, 16).NumberFormat = "0.00E+00"
        
        ' Auto-fit columns I to P
        ws.Columns("I:P").AutoFit
        
    Next ws
    
End Sub

