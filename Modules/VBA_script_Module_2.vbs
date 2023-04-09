Attribute VB_Name = "Module2"
Sub alphabetica_testingl()

Dim g As Integer
 
For Each ws In Worksheets

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

For g = 2 To LastRow

    If (ws.Cells(g, 10).Value > "0") Then


  ws.Cells(g, 10).Interior.ColorIndex = 4
  

   Else

  ws.Cells(g, 10).Interior.ColorIndex = 3

End If
   
Next g

For g = 2 To LastRow

    If (ws.Cells(g, 11).Value > "0") Then


  ws.Cells(g, 11).Interior.ColorIndex = 4
  

   Else

  ws.Cells(g, 11).Interior.ColorIndex = 3

End If
   
Next g


Next ws

End Sub
