Sub stock_data()

    Dim lastrow As Double
    Dim tablerow As Integer
    Dim stockopen As Double
    Dim stockclose As Double
    Dim ticker As String
    Dim yearly As Double
    Dim percent As Double
    Dim totalstocks As Double
    Dim max As Double
    Dim min As Double
    Dim volume As Double
    Dim ws As Worksheet
    
  For Each ws In ThisWorkbook.Worksheets
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(2, 14).Value = "Greatest % increase"
    ws.Cells(3, 14).Value = "Greatest % decrease"
    ws.Cells(4, 14).Value = "Greatest total volume"
    
    tablerow = 2
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    stockopen = ws.Cells(2, 3).Value
    
    For i = 2 To lastrow
                   
        stockclose = ws.Cells(i, 6).Value
        yearly = stockclose - stockopen
        totalstocks = totalstocks + ws.Cells(i, 7).Value
        
        If stockclose = 0 Then
                percent = 0
            Else
                percent = (stockclose - stockopen) / stockclose
        End If
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ws.Range("I" & tablerow).Value = ws.Cells(i, 1).Value
            ws.Range("J" & tablerow).Value = yearly
            ws.Range("K" & tablerow).Value = percent
            ws.Range("L" & tablerow).Value = totalstocks
            yearly = 0
            percent = 0
            totalstocks = 0
            tablerow = tablerow + 1
            stockopen = ws.Cells(i + 1, 3)
        End If
        
    Next i
    
    For j = 2 To lastrow
        If ws.Cells(j, 10).Value > 0 Then
            ws.Cells(j, 10).Interior.ColorIndex = 4
        Else
            ws.Cells(j, 10).Interior.ColorIndex = 3
        End If
        
         If ws.Cells(j, 11).Value > 0 Then
            ws.Cells(j, 11).Interior.ColorIndex = 4
        Else
            ws.Cells(j, 11).Interior.ColorIndex = 3
        End If
    Next j
    
  ws.Columns("K").NumberFormat = "0.00%"
  ws.Range("O2:O3").NumberFormat = "0.00%"
  
  max = WorksheetFunction.max(ws.Range("K2:K1000000").Value)
  ws.Range("O2") = max
  min = WorksheetFunction.min(ws.Range("K2:K1000000").Value)
  ws.Range("O3") = min
  volume = WorksheetFunction.max(ws.Range("L2:L1000000").Value)
  ws.Range("O4") = volume
  
 Next ws
  
End Sub

Sub clear()
Dim ws As Worksheet
    
  For Each ws In ThisWorkbook.Worksheets
  
    ws.Range("I1:O1000000").clear

  Next ws

End Sub