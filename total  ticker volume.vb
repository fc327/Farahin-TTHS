Sub total_volume()
 For Each ws In Worksheets


Dim WorksheetName As String

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
WorksheetName = ws.Name

Dim ticker As String

Dim volume_Total As Double
volume_Total = 0

  Dim Summary_Table_Row As Integer
 Summary_Table_Row = 2

ws.Cells(1, 10).Value = "Ticker"
ws.Cells(1, 11).Value = "Total Ticker Value"


LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row

  For i = 2 To LastRow
   If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

     ticker = ws.Cells(i, 1).Value

     volume_Total = volume_Total + Cells(i, 7).Value

     ws.Cells(Summary_Table_Row, 10).Value = ticker

     ws.Cells(Summary_Table_Row, 11).Value = volume_Total

      Summary_Table_Row = Summary_Table_Row + 1

      volume_Total = 0

 
   Else

     volume_Total = volume_Total + ws.Cells(i, 7).Value

   End If

 Next i

Next ws

End Sub
