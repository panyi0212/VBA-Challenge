Attribute VB_Name = "Module3"
Sub alphabetica_testingl()

Dim c As Integer

For Each ws In ThisWorkbook.Worksheets

'set headers

ws.Cells(2, 15).Value = "Greatest % increase"

ws.Cells(3, 15).Value = "Greatest % decrease"

ws.Cells(4, 15).Value = "Greatest total volume"

ws.Cells(1, 16).Value = "Ticker"

ws.Cells(1, 17).Value = "Value"


LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

For c = 2 To LastRow

If (ws.Cells(c, 11).Value = WorksheetFunction.Max(ws.Range("K2:K" & LastRow))) Then
 
ws.Cells(2, 16).Value = ws.Cells(c, 9).Value
ws.Cells(2, 17).Value = ws.Cells(c, 11).Value

End If

Next c

For c = 2 To LastRow

If (ws.Cells(c, 11).Value = WorksheetFunction.Min(ws.Range("K2:K" & LastRow))) Then
 
ws.Cells(3, 16).Value = ws.Cells(c, 9).Value
ws.Cells(3, 17).Value = ws.Cells(c, 11).Value

End If

Next c

For c = 2 To LastRow

If (ws.Cells(c, 12).Value = WorksheetFunction.Max(ws.Range("L2:L" & LastRow))) Then
 
ws.Cells(4, 16).Value = ws.Cells(c, 9).Value
ws.Cells(4, 17).Value = ws.Cells(c, 12).Value

End If

Next c


Next ws




End Sub
