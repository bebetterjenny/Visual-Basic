Sub cartonCounter()

'Get the number of last row
Dim lastRow As Integer
lastRow = Cells(Rows.Count, 1).End(xlUp).Row

Dim i As Integer
Dim sum As Integer
sum = 0

For lastRow = 2 To lastRow
    i = lastRow
    Range("E" & i).Select
    ActiveCell.FormulaR1C1 = _
        "=LEN(TRIM(RC[-1]))-LEN(SUBSTITUTE(TRIM(RC[-1]),"" "",""""))+1"
    Range("E" & i).Select
    sum = sum + Range("E" & i).Value
Next

Columns("E:E").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
Range("E1").Select
Application.CutCopyMode = False
Range("E" & lastRow).Value = sum
Range("E" & lastRow).Select



End Sub