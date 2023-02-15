Attribute VB_Name = "Module1"
Sub Stonks()
'Loop through each worksheet
For Each ws In Worksheets

'Defining variables and labeling columns
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Open"
ws.Range("K1").Value = "Close"
ws.Range("N1").Value = "Total Stock Volume"
Dim lastRow, i, Count As Long
Dim Volume As Double
Dim StockName, Checker As String
Count = 2
Volume = 0
lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Loop to record open, close, volume for each stock for each year
For i = 2 To lastRow
    StockName = ws.Cells(i, 1).Value
    If ws.Cells(i - 1, 1).Value <> StockName Then
        ws.Range("I" & Count).Value = StockName
        ws.Range("J" & Count).Value = ws.Cells(i, 3).Value
        Volume = ws.Cells(i, 7)
    ElseIf ws.Cells(i + 1, 1) <> StockName Then
        ws.Range("K" & Count).Value = ws.Cells(i, 6).Value
        ws.Range("N" & Count).Value = Volume
        Volume = 0
        Count = Count + 1
    Else: Volume = Volume + ws.Cells(i, 7).Value
    End If

Next i

'Loop to record change in percent over year
ws.Range("L1").Value = "Yearly Change"
ws.Range("M1").Value = "Percent Change"
Dim r As Integer
Dim lastRow2 As Long
lastRow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row

For r = 2 To lastRow2
    ws.Cells(r, 12) = ws.Cells(r, 11) - ws.Cells(r, 10)
    If ws.Cells(r, 10) <> 0 Then
        ws.Cells(r, 13) = (ws.Cells(r, 12) / ws.Cells(r, 10))
        'Formatting percent change row
        ws.Cells(r, 13).Value = FormatPercent(ws.Cells(r, 13))
    End If

    'Change cell color if % Change is + or -
    If ws.Cells(r, 12).Value < 0 Then
        ws.Cells(r, 12).Interior.ColorIndex = 3
    ElseIf ws.Cells(r, 12).Value > 0 Then
        ws.Cells(r, 12).Interior.ColorIndex = 43
    ElseIf ws.Cells(r, 12).Value = 0 Then
        ws.Cells(r, 12).Interior.ColorIndex = 7
    End If
    
    If ws.Cells(r, 13).Value < 0 Then
        ws.Cells(r, 13).Interior.ColorIndex = 3
    ElseIf ws.Cells(r, 13).Value > 0 Then
        ws.Cells(r, 13).Interior.ColorIndex = 43
    ElseIf ws.Cells(r, 13).Value = 0 Then
        ws.Cells(r, 13).Interior.ColorIndex = 7
    End If

Next r

'Find largest % increase, % decrease, total volume
ws.Range("Q2").Value = "Greatest % Increase"
ws.Range("Q3").Value = "Greatest % Decrease"
ws.Range("Q4").Value = "Greatest Total Volume"
ws.Range("R1").Value = "Ticker"
ws.Range("S1").Value = "Value"
Dim l As Integer
Dim MaxPerChange, MinPerChange, TotalVol As Double
MaxPerChange = 0
MinPerChange = 0
TotalVol = 0

'MaxPerChange = Application.WorksheetFunction.Max(Range("L:L"))
'MinPerChange = Application.WorksheetFunction.Min(Range("L:L"))
'TotalVol = Application.WorksheetFunction.Max(Range("N:N"))

'Find values, move them to the right cells, format cells
For l = 2 To lastRow2
    If ws.Cells(l, 13).Value > MaxPerChange Then
        MaxPerChange = ws.Cells(l, 13).Value
        ws.Range("S2") = MaxPerChange
        ws.Range("S2").Value = FormatPercent(ws.Range("S2"))
        ws.Range("S2").Interior.ColorIndex = 43
        ws.Range("R2").Value = ws.Cells(l, 9).Value
    End If
    If ws.Cells(l, 13).Value < MinPerChange Then
        MinPerChange = ws.Cells(l, 13).Value
        ws.Range("S3") = MinPerChange
        ws.Range("S3").Value = FormatPercent(ws.Range("S3"))
        ws.Range("S3").Interior.ColorIndex = 3
        ws.Range("R3").Value = ws.Cells(l, 9).Value
    End If
    If ws.Cells(l, 14).Value > TotalVol Then
        TotalVol = ws.Cells(l, 14).Value
        ws.Range("S4") = TotalVol
        ws.Range("R4").Value = ws.Cells(l, 9).Value
        ws.Range("S4").NumberFormat = "0"
    End If
Next l

'Bold column headings
ws.Range("A1:S1").Font.Bold = True
ws.Range("Q2:Q4").Font.Bold = True

'Autofit columns
ws.Columns("A:S").AutoFit

Next ws

End Sub


