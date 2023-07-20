Attribute VB_Name = "Module1"
Sub stockAnalysis()
    
    For Each ws In ThisWorkbook.Worksheets
        
        Dim tickerSymbols() As String
        Dim yearlyChanges() As Double
        Dim percentageChanges() As Double
        Dim totalVolumes() As Long
        Dim lastRow As Long
        Dim summaryRow As Long
    
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
        
        summaryRow = 2
        
        maxIncrease = 0
        maxDecrease = 0
        maxVolume = 0
        
        For i = 2 To lastRow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Ticker = ws.Cells(i, 1).Value
                
                openingPrice = ws.Cells(i, 3).Value
                closingPrice = ws.Cells(i, 6).Value
                
                yearlyChange = closingPrice - openingPrice
            
                
            If openingPrice <> 0 Then
                percentChange = (yearlyChange / openingPrice) * 100
            Else
                percentChange = 0
            End If
        
                
            stock = ws.Cells(i, 1).Value
    
            volume = ws.Cells(i, 7).Value
            
            ws.Cells(summaryRow, 9).Value = Ticker
            ws.Cells(summaryRow, 10).Value = yearlyChange
            ws.Cells(summaryRow, 11).Value = percentChange
            ws.Cells(summaryRow, 12).Value = Application.WorksheetFunction.Sum(ws.Cells(i - totalVolumes(summaryRow) + 1, 7), ws.Cells(i, 7))
            
            
            If (closingPrice - ws.Cells(i - 1, 6).Value) / (ws.Cells(i - 1, 6).Value) * 100 > maxIncrease Then
                maxIncrease = (closingPrice - ws.Cells(i - 1, 6).Value) / (ws.Cells(i - 1, 6).Value) * 100
                maxIncreaseStock = stock
            End If
            If (closingPrice - ws.Cells(i - 1, 6).Value) / (ws.Cells(i - 1, 6).Value) * 100 < maxDecrease Then
                maxDecrease = (closingPrice - ws.Cells(i - 1, 6).Value) / (ws.Cells(i - 1, 6).Value) * 100
                maxDecreaseStock = stock
            End If

            If volume > maxVolume Then
                maxVolume = volume
                maxVolumeStock = stock
            End If
         
            If yearlyChange < 0 Then
                ws.Cells(summaryRow, 10).Interior.ColorIndex = 3
            ElseIf yearlyChange > 0 Then
                ws.Cells(summaryRow, 10).Interior.ColorIndex = 4
            End If
                     
            summaryRow = summaryRow + 1
            
            
    End If
            
            
        totalVolumes(summaryRow) = totalVolumes(summaryRow) + volume
            
         
         Next i
         
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("Q2").Value = maxIncreaseStock
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").Value = maxDecreaseStock
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Range("Q4").Value = maxVolumeStock
        ws.Range("Q4").NumberFormat = "#,##0"
    
    Next ws

End Sub
