Sub imgNamer()
Dim lastRow As Integer
lastRow = Cells(Rows.Count, 1).End(xlUp).Row
Dim i As Integer

For i = 2 To lastRow
    Range("B" & i).Select
    ActiveCell.FormulaR1C1 = "=""/""&RC[-1]&""-1.jpg"""
    Range("C" & i).Select
    ActiveCell.FormulaR1C1 = _
        "=""/""&RC[-2]&""-1.jpg,/""&RC[-2]&""-2.jpg,/""&RC[-2]&""-3.jpg,/""&RC[-2]&""-4.jpg"""

Next

Cells.Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
Application.CutCopyMode = False

End Sub