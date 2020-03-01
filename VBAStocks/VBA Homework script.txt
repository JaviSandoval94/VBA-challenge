Sub stockTotal()
    Dim lastRow As Long
    Dim runningRow As Long
    Dim tickerSymbol As String
    Dim stockOpen As Double
    Dim stockClose As Double
    Dim stockVolume As Double
    Dim maxInc As Double
    Dim maxDec As Double
    Dim maxVol As Double
    Dim maxIncTck As String
    Dim maxDecTck As String
    Dim maxVolTck As String
        
    'Iterate this code through each worksheet in the workbook. Deifine the last row of data each time.
    For Each ws In Worksheets
    
        lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row

        'Start variables at initial values.
        runningRow = 2
        maxInc = 0
        maxDec = 0
        maxVol = 0
        
        'Sort rows to run the rest of the code: alphabetically depending on the ticker key and then by date.
        ws.Columns("A:G").Sort key1:=ws.Range("A2"), order1:=xlAscending, Header:=xlYes, key2:=ws.Range("B2"), order2:=xlAscending, Header:=xlYes
    
        'Generate the summary table headers.
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        'Define opening price of the first ticker symbol.
        stockOpen = ws.Cells(2, 3).Value
    
        'Sum values of total stock volume for all the entries of each ticker symbol value.
        For i = 2 To lastRow
            tickerSymbol = ws.Cells(i, 1).Value
            stockVolume = stockVolume + ws.Cells(i, 7).Value
        
            'Before running into the next ticker symbol value, calculate the total stock price difference and save the values in the summary table.
            If ws.Cells(i + 1, 1) <> tickerSymbol Then
                stockClose = ws.Cells(i, 6).Value
                ws.Cells(runningRow, 9).Value = tickerSymbol
                ws.Cells(runningRow, 10).Value = stockClose - stockOpen
                
                'Consider the case when the opening price is 0, which will lead to an error when calculating the percentage value.
                If stockOpen = 0 Then
                    'If the closing price is also 0, assign 0% to the total percentage stock value change.
                    If stockClose = 0 Then
                        ws.Cells(runningRow, 11).Value = 0
                    'If the closing price is different to 0, assign 100% difference (consider whether the closing stock value is positive or negative).
                    Else
                        ws.Cells(runningRow, 11).Value = stockClose / Abs(stockClose)
                    End If
                'If the total opening price is different to 0, calculate percentage change normally and save it in the summary table.
                Else
                    ws.Cells(runningRow, 11).Value = stockClose / stockOpen - 1
                End If
                'Save the accumulated stock volume value in the summary table.
                ws.Cells(runningRow, 12).Value = stockVolume
            
                'For each ticker symbol, determine whether or not the volume value at each row is the maximum. Save the maximum value in the variables.
                If ws.Cells(runningRow, 12).Value > maxVol Then
                    maxVol = ws.Cells(runningRow, 12).Value
                    maxVolTck = tickerSymbol
                End If
                
                'For each data row, determine whether or not the change value at each row is the greatest increase. If it is, save the values in the variables.
                If ws.Cells(runningRow, 11).Value > maxInc Then
                    maxInc = ws.Cells(runningRow, 11).Value
                    maxIncTck = tickerSymbol
                    
                'For each ticker symbol, determine whether or not the change value at each row is the greatest decrease. If it is, save the values in the variables.
                ElseIf ws.Cells(runningRow, 11).Value < maxDec Then
                    maxDec = ws.Cells(runningRow, 11).Value
                    maxDecTck = tickerSymbol
                End If
            
                'Format cells at the summary table.
                ws.Cells(runningRow, 11).NumberFormat = "0.00%"
                ws.Cells(runningRow, 12).NumberFormat = "#,###"
            
                If ws.Cells(runningRow, 10) < 0 Then
                    ws.Cells(runningRow, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(runningRow, 10).Interior.ColorIndex = 4
                End If
                
                'When the next ticker value is different, move one row down in the summary table and restart variables to 0.
                runningRow = runningRow + 1
                stockOpen = ws.Cells(i + 1, 3).Value
                stockClose = 0
                stockVolume = 0
        
            End If
        
        Next i
    
        'Generate the max/min summary table headers.
        ws.Cells(2, 15).Value = "Greatest % increase"
        ws.Cells(3, 15).Value = "Greatest % decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
    
        'Autofit the values at column O to the written text.
        ws.Columns("O").AutoFit
    
        'Save the greatest volume, greatest increase and greatest decrese in the max/min summary table. Format cells accordingly and autofit the volume column.
        ws.Cells(2, 16).Value = maxIncTck
        ws.Cells(2, 17).Value = maxInc
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 16).Value = maxDecTck
        ws.Cells(3, 17).Value = maxDec
        ws.Cells(3, 17).NumberFormat = "0.00%"
        ws.Cells(4, 16).Value = maxVolTck
        ws.Cells(4, 17).Value = maxVol
        ws.Cells(4, 17).NumberFormat = "#,###"
        ws.Columns("Q").AutoFit
                
    Next ws
    
    'Display a message when the summary tables are ready in every datasheet.
    MsgBox ("Your data is ready!")
End Sub






