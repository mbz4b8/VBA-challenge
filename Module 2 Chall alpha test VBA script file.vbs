Attribute VB_Name = "Module1"
Sub Mdule():

Dim i As Long
Dim oi As Long
Dim percChange As Double
Dim yearChange As Double
Dim totalvolume As Double
Dim LastRow As Long
Dim si As Integer
Dim ticker As String
Dim maxvalue As Double
Dim greatincrease As Double
Dim greatdecrease As Double

For Each ws In ThisWorkbook.Sheets

si = 2
oi = 2
totalvolume = 2
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"


For i = 2 To LastRow

    totalvolume = totalvolume + ws.Cells(i, "G").Value
    If ws.Cells(i, "A").Value <> ws.Cells(i + 1, "A").Value Then
        yearChange = ws.Cells(i, "F").Value - ws.Cells(oi, "C").Value
        percChange = yearChange / ws.Cells(oi, "C").Value
        ticker = ws.Cells(i, "A").Value
        
        ws.Range("i" & si).Value = ticker
        ws.Range("j" & si).Value = yearChange
        ws.Range("k" & si).Value = percChange
        ws.Range("k" & si).NumberFormat = "0.00%"
         If ws.Range("k" & si).Value >= 0 Then
        ws.Range("k" & si).Interior.ColorIndex = 4
        ElseIf ws.Range("k" & si).Value < 0 Then
        ws.Range("k" & si).Interior.ColorIndex = 3
        Else: ws.Range("k" & si).Interior.ColorIndex = xlNone
        End If
        ws.Range("l" & si).Value = totalvolume
        If ws.Range("J" & si).Value >= 0 Then
        ws.Range("J" & si).Interior.ColorIndex = 4
        ElseIf ws.Range("J" & si).Value < 0 Then
        ws.Range("J" & si).Interior.ColorIndex = 3
        Else: ws.Range("J" & si).Interior.ColorIndex = xlNone
        End If
        
        
        oi = i + 1
        si = si + 1
        totalvolume = 0
        
    End If

Next i

maxvalue = Application.WorksheetFunction.Max(ws.Columns("L"))
greatincrease = Application.WorksheetFunction.Max(ws.Columns("K"))
greatdecrease = Application.WorksheetFunction.Min(ws.Columns("K"))

ws.Cells(4, 17).Value = maxvalue
ws.Cells(2, 17).Value = greatincrease
ws.Cells(2, 17).NumberFormat = "0.00%"
ws.Cells(3, 17).Value = greatdecrease
ws.Cells(3, 17).NumberFormat = "0.00"

For i = 2 To LastRow

If ws.Cells(i, "l") = maxvalue Then
ws.Cells(4, 16).Value = ws.Cells(i, "I")
End If

If ws.Cells(i, "K") = greatincrease Then
ws.Cells(2, 16).Value = ws.Cells(i, "I")
End If

If ws.Cells(i, "K") = greatdecrease Then
ws.Cells(3, 16).Value = ws.Cells(i, "I")
End If

Next i

Next ws

End Sub


