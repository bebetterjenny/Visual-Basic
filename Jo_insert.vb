Sub Jo_insert_air()

    ActiveCell.Rows("1:1").EntireRow.Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveCell.FormulaR1C1 = "=R[1]C[3] & CHAR(10) & R[1]C[2]"
    ActiveCell.Offset(0, 0).Range("A1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    ActiveCell.Offset(0, 0).Range("A1").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    ActiveCell.Offset(0, 1).Range("A1").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    ActiveCell.Offset(2, -1).Range("A1").Select
    

End Sub
Sub Jo_insert_sea()

    ActiveCell.Rows("1:1").EntireRow.Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveCell.FormulaR1C1 = "=R[1]C[3] & CHAR(10) & R[1]C[2]"
    ActiveCell.Offset(0, 0).Range("A1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    ActiveCell.Offset(0, 0).Range("A1").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    ActiveCell.Offset(0, 1).Range("A1").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    ActiveCell.Offset(2, -1).Range("A1").Select
    

End Sub
