Sub insertCopy()

ActiveCell.Rows("1:1").EntireRow.Select
Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
ActiveCell.FormulaR1C1 = "=R[1]C[0]"
ActiveCell.Offset(0, 0).Range("A1").Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
Application.CutCopyMode = False

End Sub