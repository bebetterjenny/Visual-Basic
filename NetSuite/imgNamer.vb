Sub imgNamer()


'Dim lastRow As Integer
'lastRow = Cells(Rows.Count, 1).End(xlUp).Row
'MsgBox (lastRow)
Dim i As Integer

For i = 30001 To 35000
    If Range("A" & i).Value = "" Then
        Range("A" & i).Value = Range("A" & i - 1).Value
    End If
Next


For i = 30001 To 35000

        Sheets("Sheet1").Select
        Range("B" & i).Select
        ActiveCell.FormulaR1C1 = "=""/""&RC[-1]&""-"""


Next

Cells.Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
Application.CutCopyMode = False


End Sub

Sub imgNumber()

Dim lastRow As Integer
lastRow = Cells(Rows.Count, 1).End(xlUp).Row
Dim i As Integer
Dim j As Integer

For i = 2 To lastRow

        Range("C" & i).Value = (i + 3) Mod 4

Next


End Sub

Sub rmvSCP()

Dim lastRow As Integer
lastRow = Cells(Rows.Count, 1).End(xlUp).Row
Dim i As Integer

For i = 2 To lastRow
    Range("A" & i).Select
    If (InStr(1, Range("A" & i), "SCP") = 1) Then
        If InStr(Range("B" & i), "-4.jpg") Then
            'MsgBox (i)
            Rows(i & ":" & i).Select
            Selection.Delete Shift:=xlUp
        End If
    End If

Next



    'Rows("33666:33666").Select
    'Selection.Delete Shift:=xlUp
End Sub

