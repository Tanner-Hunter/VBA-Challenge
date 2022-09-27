Sub VBA_homework()
'Creat loop for all worksheets

For Each ws In Worksheets


'declare variables
Dim Summary_Table_Row As Integer
Dim Opening, percent, high, low, closing, volume As Double
Dim ticker As String
Dim change As Double

'setting initial values.

Summary_Table_Row = 2
Opening = ws.Cells(2, 3).Value
volume = 0

'create an i
For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row

'Creating if statement.
If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
volume = volume + ws.Cells(i, 7).Value

ElseIf ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
volume = volume + ws.Cells(i, 7).Value

ticker = ws.Cells(i, 1).Value
closing = ws.Cells(i, 6).Value
change = Opening - closing
percent = (change / Opening) * 100

ws.Range("I" & Summary_Table_Row).Value = ticker
ws.Range("J" & Summary_Table_Row).Value = change
ws.Range("K" & Summary_Table_Row).Value = percent
ws.Range("L" & Summary_Table_Row).Value = volume

'Reseting values.
Summary_Table_Row = Summary_Table_Row + 1
volume = 0


End If
Opening = ws.Cells(i + 1, 3).Value
Next i

'Color yearly change based on psitive or negative change.

For c = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
If ws.Cells(c, 10).Value > 0 Then
ws.Cells(c, 10).Interior.ColorIndex = 4
ElseIf ws.Cells(c, 10).Value < 0 Then
ws.Cells(c, 10).Interior.ColorIndex = 3

End If
Next c

Next ws
End Sub
