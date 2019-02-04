Sub unit2Homework()
' variables
Dim lastRow As Long
Dim i As Long
Dim total As Double
Dim tickerCount As Integer
Dim ws As Worksheet

' loop through each worksheet
For Each ws In Worksheets
    ws.Activate

    ' use End(xlUp) to determine Last Row with Data, in one column (column B)
    lastRow = ActiveSheet.Range("B" & Rows.Count).End(xlUp).Row
    
    ' name the new column headers
    ActiveSheet.Cells(1, 9).Value = "Ticker"
    ActiveSheet.Cells(1, 10).Value = "Total Stock Volume"
    
    ' set couneter to keep sum up each ticker
    tickerCount = 1
    
    ' loop through rows to calculate the volum for each uniq ticker symbol
    For i = 2 To lastRow
        
        If ActiveSheet.Range("A" & i).Value = ActiveSheet.Range("A" & i + 1).Value Then
            total = total + ActiveSheet.Cells(i, 7).Value
        Else
            tickerCount = tickerCount + 1
            total = total + Cells(i, 7).Value
            ActiveSheet.Cells(tickerCount, 9).Value = ActiveSheet.Range("A" & i).Value
            ActiveSheet.Cells(tickerCount, 10).Value = total
            total = 0
        End If
    
    Next i

Next ws

End Sub
