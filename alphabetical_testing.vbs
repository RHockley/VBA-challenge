Sub tickerChecker()

  'ticker code and volume count appear twice
  'still need total price and percentage change
  'not working across multiple sheets
  
Dim ws As Worksheet
Dim tickerCode As String
Dim volumeCount As Double
volumeCount = 0
Dim summaryTableRow As Integer
summaryTableRow = 2


For Each ws In Worksheets

            Range("I1").Value = "Ticker"
            Range("J1").Value = "Price Change"
            Range("K1").Value = "Percentage Change"
            Range("L1").Value = "Total Stock Volume"
            Columns("A:M").AutoFit
            
    For i = 2 To 22771
    
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            tickerCode = Cells(i, 1).Value
            volumeCount = volumeCount + Cells(i, 7).Value
            
            Range("I" & summaryTableRow).Value = tickerCode
            Range("L" & summaryTableRow).Value = volumeCount
            
            summaryTableRow = summaryTableRow + 1
            
            volumeCount = 0
            
            Else
            
                volumeCount = volumeCount + Cells(i, 7).Value
                
            End If
            
        Next i
        
    
Next ws

End Sub
