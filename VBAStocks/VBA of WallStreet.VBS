Attribute VB_Name = "Module1"
Option Explicit

Sub wallStreet()

Dim ws As Worksheet
    For Each ws In Worksheets
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(2, 14).Value = "Greatest % Increase"
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    ws.Cells(4, 14).Value = "Greatest Total Volume"
    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value"
    
    Dim openPrice, closePrice As Double
    Dim yearChange, percChange, volume, row, lastRow, i As Long
    Dim ticker As String
    row = 2
    volume = 0
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).row
    
    ' Set Initial Open Price
    openPrice = ws.Cells(2, 3).Value
    ' Loop through all ticker symbols
        For i = 2 To lastRow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' Set Ticker name
                ticker = ws.Cells(i, 1).Value
                ws.Range("I" & row).Value = ticker
                ' Close Price
                closePrice = ws.Cells(i, 6).Value
                ' Yearly Change
                yearChange = closePrice - openPrice
                ws.Cells(row, 10).Value = yearChange
                If ws.Cells(row, 10).Value >= 0 Then
                    ws.Cells(row, 10).Interior.ColorIndex = 43
                Else
                    ws.Cells(row, 10).Interior.ColorIndex = 53
                End If
                'Percent Change
                If openPrice = 0 Then 'CHECK IF NEEDED ON FINAL RUN
                    percChange = 0
                Else
                    percChange = yearChange / openPrice
                    ws.Cells(row, 11).Value = percChange
                    ws.Cells(row, 11).NumberFormat = "0.00%"
                End If
                ' Add volume, add table row, set next price, reset volume
                volume = volume + ws.Cells(i, 7).Value
                ws.Cells(row, 12).Value = volume
                row = row + 1
                openPrice = ws.Cells(i + 1, 3)
                volume = 0
            'finally, if same ticker, still add volume
            Else
                volume = volume + ws.Cells(i, 7).Value
            End If
        Next i
    
        'Bonus
        For i = 2 To lastRow
            If ws.Range("K" & i).Value > ws.Range("P2").Value Then
                ws.Range("P2").Value = ws.Range("K" & i).Value
                ws.Range("O2").Value = ws.Range("I" & i).Value
            End If
        
            If ws.Range("K" & i).Value < ws.Range("P3").Value Then
                ws.Range("P3").Value = ws.Range("K" & i).Value
                ws.Range("O3").Value = ws.Range("I" & i).Value
            End If
        
            If ws.Range("L" & i).Value > ws.Range("P4").Value Then
                ws.Range("P4").Value = ws.Range("L" & i).Value
                ws.Range("O4").Value = ws.Range("I" & i).Value
            End If
        Next i
        ws.Range("P2:P3").NumberFormat = "0.00%"
            
        'Couple finishing touches
        ws.Rows("1").Font.Bold = True
        ws.Range("A:L").AutoFilter
        ws.Range("N2:P4").Interior.ColorIndex = 34
        ws.Range("N1:P1").Interior.ColorIndex = 33
        ws.UsedRange.HorizontalAlignment = xlLeft
        ws.Columns("A:R").AutoFit
    Next ws
End Sub


