Sub Stock()
'Activate All sheets
Dim ws As Worksheet
Dim starting_ws As Worksheet
Set starting_ws = ActiveSheet

For Each ws In ThisWorkbook.Worksheets
    ws.Activate
Dim ti As String
Dim total As Double
total = 0
Dim SumRow As Integer
SumRow = 2
Dim YrCh As Double
Dim PrCh As Double

'Challenge variables

Dim Mi As Double
Dim Ma As Double
Dim Most As Double

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

wscount = ActiveWorkbook.Worksheets.Count
    
        For i = 2 To LastRow
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                ti = Cells(i, 1).Value
                total = total + Cells(i, 7).Value
                
                YrCh = Cells(i, 6).Value - Cells(2, 3).Value
                PrCh = YrCh / Cells(2, 3).Value
                
                Range("I" & SumRow).Value = ti
                Range("J" & SumRow).Value = YrCh
                Range("K" & SumRow).Value = PrCh
                Range("L" & SumRow).Value = total
                
                If Range("J" & SumRow).Value < 0 Then
                    Range("J" & SumRow).Interior.ColorIndex = 3
                ElseIf Range("J" & SumRow).Value > 0 Then
                    Range("J" & SumRow).Interior.ColorIndex = 37
                End If
                
                Range("K" & SumRow).NumberFormat = "0.00%"
                
                SumRow = SumRow + 1
                total = 0
            Else
                total = total + Cells(i, 7).Value
            End If
                  
        Next i
  
   Cells(2, 15).Value = "Greatest % Increase"
   Cells(3, 15).Value = "Greatest % Decrease"
   Cells(4, 15).Value = "Greatest total volume"
   
    Ma = Application.WorksheetFunction.Max(Range("K:K"))
            Cells(2, 17).Value = Ma
            Cells(2, 17).NumberFormat = "0.00%"
    Mi = Application.WorksheetFunction.Min(Range("K:K"))
            Cells(3, 17).Value = Mi
            Cells(3, 17).NumberFormat = "0.00%"
    Most = Application.WorksheetFunction.Max(Range("L:L"))
            Cells(4, 17).Value = Most
            
    Maname = WorksheetFunction.Match(WorksheetFunction.Max(Range("K:K")), Range("K:K"), 0)
        Cells(2, 16).Value = Range("I" & Maname).Value
    Miname = WorksheetFunction.Match(WorksheetFunction.Min(Range("K:K")), Range("K:K"), 0)
        Cells(3, 16).Value = Range("I" & Miname).Value
    Moname = WorksheetFunction.Match(WorksheetFunction.Max(Range("L:L")), Range("L:L"), 0)
        Cells(4, 16).Value = Range("I" & Moname).Value
    
Next ws

starting_ws.Activate

End Sub