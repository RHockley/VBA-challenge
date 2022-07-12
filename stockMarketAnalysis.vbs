Sub stockMarketAnalysis()

Dim ticker As String
Dim yearOpen As String
Dim yearClose As String
Dim yearlyChange As Double
Dim percentageChange As Double
Dim totalStockVolume As Double
Dim startData As Integer
Dim ws As Worksheet

For Each ws In Worksheets

    ws.Range("I:L").EntireColumn.AutoFit
    ws.Range("I1").Value = "Ticker Code"
    ws.Range("J1").Value = "Total Yearly Change"
    ws.Range("K1").Value = "Total Percentage Change"
    ws.Range("L1").Value = "Total Stock Volume"

startData = 2
previousI = 1
totalStockVolume = 0

EndRow = ws.Cells(Rows.Count, "A").End(xlUp).Row

    For i = 2 To EndRow
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            ticker = ws.Cells(i, 1).Value
            
            previousI = previousI + 1
            
            yearOpen = ws.Cells(previousI, 3).Value
            yearClose = ws.Cells(i, 6).Value
            
                For j = previousI To i
                
                    totalStockVolume = totalStockVolume + ws.Cells(j, 7).Value
                    
                Next j
                
                If yearOpen = 0 Then
                
                    percentageChange = yearClose
                    
                Else
                    yearlyChange = yearClose - yearOpen
                    percentageChange = yearlyChange / yearOpen
                    
                End If
                
     
        
        ws.Cells(startData, 9).Value = ticker
        ws.Cells(startData, 10).Value = yearlyChange
        ws.Cells(startData, 11).Value = percentageChange
        
        ws.Cells(startData, 11).NumberFormat = "0.00%"
        ws.Cells(startData, 12).Value = totalStockVolume
        
        startData = startData + 1
        
        totalStockVolume = 0
        yearlyChange = 0
        percentageChange = 0
        
        previousI = i
        
        End If
        
        Next i

    jEndRow = ws.Cells(Rows.Count, "J").End(xlUp).Row
        For j = 2 To jEndRow
            If ws.Cells(j, 10) > 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 4

            Else
                ws.Cells(j, 10).Interior.ColorIndex = 3
            End If

        Next j

Next ws

End Sub


