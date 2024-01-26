Attribute VB_Name = "Final2"

Sub StockAnalysis()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim maxIncrease As Double
    Dim maxDecrease As Double
    Dim maxVolume As Double
    Dim maxIncreaseTicker As String
    Dim maxDecreaseTicker As String
    Dim maxVolumeTicker As String
    
    ' Loop through each worksheet
    For Each ws In ThisWorkbook.Sheets
        ' Initialize variables for each worksheet
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        maxIncrease = 0
        maxDecrease = 0
        maxVolume = 0
        maxIncreaseTicker = ""
        maxDecreaseTicker = ""
        maxVolumeTicker = ""
        
        ' Add headers for the output columns
        ws.Cells(1, 10).Value = "Ticker"
        ws.Cells(1, 11).Value = "Yearly Change"
        ws.Cells(1, 12).Value = "Percent Change"
        ws.Cells(1, 13).Value = "Stock Volume"
        
        ' Loop through each row in the worksheet
        For i = 2 To lastRow
            ' Get values from the row
            ticker = ws.Cells(i, 1).Value
            openPrice = ws.Cells(i, 3).Value
            closePrice = ws.Cells(i, 6).Value
            totalVolume = ws.Cells(i, 7).Value
            
            ' Calculate yearly change and percent change
            yearlyChange = closePrice - openPrice
            If openPrice <> 0 Then
                percentChange = (yearlyChange / openPrice) * 100
            Else
                percentChange = 0
            End If
            
            ' Output values to new columns
            ws.Cells(i, 10).Value = ticker
            ws.Cells(i, 11).Value = yearlyChange
            ws.Cells(i, 12).Value = percentChange
            ws.Cells(i, 13).Value = totalVolume
            
            ' Apply conditional formatting for yearly change and percent change
            If yearlyChange > 0 Then
                ws.Cells(i, 11).Interior.Color = RGB(0, 255, 0) ' Green for positive change
            ElseIf yearlyChange < 0 Then
                ws.Cells(i, 11).Interior.Color = RGB(255, 0, 0) ' Red for negative change
            End If
            
            If percentChange > 0 Then
                ws.Cells(i, 12).Interior.Color = RGB(0, 255, 0) ' Green for positive change
            ElseIf percentChange < 0 Then
                ws.Cells(i, 12).Interior.Color = RGB(255, 0, 0) ' Red for negative change
            End If
            
            
            ' Update maximum values
            If percentChange > maxIncrease Then
                maxIncrease = percentChange
                maxIncreaseTicker = ticker
            ElseIf percentChange < maxDecrease Then
                maxDecrease = percentChange
                maxDecreaseTicker = ticker
            End If
            
            If totalVolume > maxVolume Then
                maxVolume = totalVolume
                maxVolumeTicker = ticker
            End If
        Next i

        
        ' Output greatest % increase, % decrease, and total volume
        ws.Cells(1, 16).Value = "Greatest % Increase"
        ws.Cells(2, 16).Value = "Greatest % Decrease"
        ws.Cells(3, 16).Value = "Greatest Total Volume"
        ws.Cells(1, 17).Value = maxIncreaseTicker
        ws.Cells(2, 17).Value = maxDecreaseTicker
        ws.Cells(3, 17).Value = maxVolumeTicker
        ws.Cells(1, 18).Value = maxIncrease
        ws.Cells(2, 18).Value = maxDecrease
        ws.Cells(3, 18).Value = maxVolume
   
    Next ws
End Sub

