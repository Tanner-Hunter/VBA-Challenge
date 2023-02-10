Attribute VB_Name = "Module1"

Sub VBA_challenge()


'define variables

Dim ws As Worksheet
For Each ws In Worksheets
On Error Resume Next
Dim ticker As String
Dim volume As Double
Dim year_open As Double
Dim year_closed As Double
Dim yearly_change As Double
Dim percentage_change As Double
Dim Summary_Table_Row As Integer
Dim x As Double
x = 2

Summary_Table_Row = 2

Dim LRow As Long

LRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  ws.Cells(1, 9).Value = "Ticker"
 ws.Cells(1, 10).Value = "Yearly_change"
  ws.Cells(1, 11).Value = "Percent_change"
   ws.Cells(1, 12).Value = "Total_volume"
  
   'create for loop
    For i = 2 To LRow
    
    
'Check if we are still within the same ticker, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
    'output the following values
    
   ticker = ws.Cells(i, 1).Value
volume = volume + ws.Cells(i, 7).Value
yearly_change = ws.Cells(i, 6).Value - ws.Cells(x, 3).Value
percent_change = yearly_change / ws.Cells(x, 3).Value
ws.Cells(Summary_Table_Row, 9).Value = ticker
ws.Cells(Summary_Table_Row, 10).Value = yearly_change
ws.Cells(Summary_Table_Row, 10).NumberFormat = "0.00"
ws.Cells(Summary_Table_Row, 11).Value = percent_change
ws.Cells(Summary_Table_Row, 11).NumberFormat = "0.00%"
ws.Cells(Summary_Table_Row, 12).Value = volume
    
 If ws.Cells(Summary_Table_Row, 10).Value > 0 Then
ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
    Else
ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
    End If
    
    'add one to the summary_table_row
    
Summary_Table_Row = Summary_Table_Row + 1
             
'reset the volume
             volume = 0
             x = i + 1
             
' If the cell following a row is the same ticker
Else
volume = volume + ws.Cells(i, 7).Value

         

     End If


    Next i
    
    Next ws

End Sub
