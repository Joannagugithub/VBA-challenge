Attribute VB_Name = "Module1"
Sub vbahomework()


Dim ws As Worksheet
For Each ws In Worksheets

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 9).Font.Bold = True

ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 10).Font.Bold = True

ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 11).Font.Bold = True

ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(1, 12).Font.Bold = True

ws.Cells(1, 17).Value = "Ticker"
ws.Cells(1, 17).Font.Bold = True

ws.Cells(1, 18).Value = "value"
ws.Cells(1, 18).Font.Bold = True

ws.Cells(2, 16).Value = "Greatest % Increase"
ws.Cells(2, 16).Font.Bold = True

ws.Cells(3, 16).Value = "Greatest % Decrease"
ws.Cells(3, 16).Font.Bold = True

ws.Cells(4, 16).Value = "Greatest total volumn"
ws.Cells(4, 16).Font.Bold = True




Dim totalvolumn As Double
totalvolumn = 0

Dim summary_table_row As Integer
summary_table_row = 2

Dim columnnumber As Double
Dim reviewer As String

Dim tickernumber As Double
tickernumber = 0


Dim rowcount1 As Double
rowcount1 = ws.Range("A2", ws.Range("A2").End(xlDown)).Rows.Count

For i = 2 To rowcount1
   
 
If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
   
   tickername = ws.Cells(i, 1).Value

   
   totalvolumn = totalvolumn + ws.Cells(i, 7).Value
   
   
   ws.Range("I" & summary_table_row).Value = tickername
   ws.Range("L" & summary_table_row).Value = totalvolumn
   
   
   ws.Range("J" & summary_table_row).Value = ws.Cells(i, 6).Value - ws.Cells(i - tickernumber, 3).Value
   
   If ws.Cells(i - tickernumber, 3).Value <> 0 Then
   ws.Range("k" & summary_table_row).Value = (ws.Cells(i, 6).Value - ws.Cells(i - tickernumber, 3).Value) / ws.Cells(i - tickernumber, 3).Value
   Else:
   ws.Range("k" & summary_table_row).Value = 0
   End If
   
   
   ws.Range("k" & summary_table_row).NumberFormat = "0.00%"
 
   
   If ws.Range("J" & summary_table_row).Value < 0 Then
   ws.Range("J" & summary_table_row).Interior.ColorIndex = 3
   
   Else:
   
   ws.Range("J" & summary_table_row).Interior.ColorIndex = 4
   End If
   
   
   
   
   
summary_table_row = summary_table_row + 1
   
    totalvolumn = 0
    tickernumber = 0
   
Else:
   
   
      totalvolumn = totalvolumn + ws.Cells(i, 7).Value
      tickernumber = tickernumber + 1
     
End If

Next i


  'homework chanllenge
 
 

Dim columnk As Range
Set columnk = ws.Range("K:K")


Dim M_increase As Double
M_increase = Application.WorksheetFunction.Max(columnk)
Dim M_decrease As Double
M_decrease = Application.WorksheetFunction.Min(columnk)

ws.Cells(3, 18).Value = M_decrease
ws.Cells(3, 18).NumberFormat = "0.00%"


ws.Cells(2, 18).Value = M_increase
ws.Cells(2, 18).NumberFormat = "0.00%"


Dim columnL As Range
Set columnL = ws.Range("L:L")

Dim G_total As Double
G_total = Application.WorksheetFunction.Max(columnL)
ws.Cells(4, 18).Value = G_total



reviewer1 = ws.Cells(2, 18).Value
reviewer2 = ws.Cells(3, 18).Value
reviewer3 = ws.Cells(4, 18).Value


Dim rowcount2 As Double
rowcount2 = ws.Range("K2", ws.Range("K2").End(xlDown)).Rows.Count

For j = 2 To rowcount2

If ws.Cells(j, 11).Value = reviewer1 Then
ws.Cells(2, 17).Value = ws.Cells(j, 9).Value
End If


If ws.Cells(j, 11).Value = reviewer2 Then
ws.Cells(3, 17).Value = ws.Cells(j, 9).Value
End If

If ws.Cells(j, 12).Value = reviewer3 Then
ws.Cells(4, 17).Value = ws.Cells(j, 9).Value
End If



Next j



Next ws

End Sub

