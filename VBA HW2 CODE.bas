Attribute VB_Name = "Module1"
Sub StockAnaysis()
    
    For Each ws In ThisWorkbook.Worksheets
        
        Dim tickerSymbols() As String
        Dim yearlyChanges() As Double
        Dim percentageChanges() As Double
        Dim totalVolumes() As Long
        Dim lastRow As Long
        Dim summary As Long
    
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        Dim maxIncrease As Double
        Dim maxDecrease As Double
        Dim maxVolume As Double
        Dim maxIncreaseStock As String
        Dim maxDecreaseStock As String
        Dim maxVolumeStock As String
        
    
        
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        SummaryRow = 2
        
        maxIncrease = 0
        maxDecrease = 0
        maxVolume = 0
        
        For i = 2 To lastRow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Ticker = ws.Cells(i, 1).Value
                
                openingPrice = ws.Cells(i, 3).Value
                
            
                'closingPrice = ws.Cells(i, 6).Value
                
                yearlyChange = closingPrice - openingPrice
            
                
            If openingPrice <> 0 Then
                percentChange = (yearlyChange / openingPrice) * 100
            Else
                percentChange = 0
            End If
        
                
            stock = ws.Cells(i, 1).Value
            closingPrice = ws.Cells(i, 5).Value
            volume = ws.Cells(i, 6).Value
            
            ws.Cells(SummaryRow, 9).Value = Ticker
            ws.Cells(SummaryRow, 10).Value = yearlyChange
            ws.Cells(SummaryRow, 11).Value = percentChange
            ws.Cells(SummaryRow, 12).Value = Application.WorksheetFunction.Sum(ws.Cells(i - totalVolume + 1, 7), ws.Cells(i, 7))
            
            
            If (closingPrice - ws.Cells(i - 1, 5).Value) / (ws.Cells(i - 1, 5).Value) * 100 > maxIncrease Then
                maxIncrease = (closingPrice - ws.Cells(i - 1, 5).Value) / (ws.Cells(i - 1, 5).Value) * 100
                maxIncreaseStock = stock
            End If
            If (closingPrice - ws.Cells(i - 1, 5).Value) / (ws.Cells(i - 1, 5).Value) * 100 < maxDecrease Then
                maxDecrease = (closingPrice - ws.Cells(i - 1, 5).Value) / (ws.Cells(i - 1, 5).Value) * 100
                maxDecreaseStock = stock
            End If

            If volume > maxVolume Then
                maxVolume = volume
                maxVolumeStock = stock
            End If
         
            If yearlyChange < 0 Then
                ws.Cells(SummaryRow, 10).Interior.ColorIndex = 3
            ElseIf yearlyChange > 0 Then
                ws.Cells(SummaryRow, 10).Interior.ColorIndex = 4
            End If
                     
    End If
            totalVolume = 0
            
            totalVolume = totalVolume + ws.Cells(i, 7).Value
            SummaryLastRow = ws.Cells(Row.Count, 9).End(xlDown).Row
         
         Next i
    Next ws

End Sub
