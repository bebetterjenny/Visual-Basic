Sub warehouseBin()

Dim lastRow As Integer
lastRow = Cells(Rows.Count, 1).End(xlUp).Row
Dim i As Integer


For i = 2 To lastRow
    Range("A" & i+1).Select
    If Range("A" & i + 1) = "" And Range("G" & i + 1) = Range("G" & i) Then
        Range("A" & i + 1) = Range("A" & i) 
    ElseIf (Range("A" & i + 1) = Range("A" & i)) Or (Range("G" & i + 1) = Range("G" & i)) Then
        'MsgBox ("对了")
        ElseIf (Range("A" & i + 1) <> Range("A" & i)) And (Range("G" & i + 1) = Range("G" & i)) Then
            MsgBox ("错了")
    End If
Next



End Sub
