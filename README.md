# VBA-challenge
Sub Stock()

Dim ws As Worksheet
For Each ws In Worksheets

   Dim ti As String
   Dim total As Double
   total = 0
   Dim SumRow As Integer
   SumRow = 2
   LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
   
    For i = 2 To LastRow
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            ti = Cells(i, 1).Value
            total = total + Cells(i, 7).Value
            Range("I" & SumRow).Value = ti
            Range("J" & SumRow).Value = total
            SumRow = SumRow + 1
            total = 0
        Else
            total = total + Cells(i, 7).Value
        End If
    Next i

Next ws
    

End Sub
