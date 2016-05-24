Sub commonUse()


'Add a new sheet
'Change SHEETNAME
Sheets.Add(After:=Sheets(Sheets.Count)).Name = "SHEETNAME"

'Get the number of last row
Dim lastRow As Integer
lastRow = Cells(Rows.Count, 1).End(xlUp).Row

'Get the number of last column
Dim lastColNum As Integer
lastColNum = Cells(1, Columns.Count).End(xlToLeft).Column

'Get the letter of last column
'Do not forget the function: ConvertToLetter
Dim lastColNum As Integer
Dim lastCol As String
lastColNum = Cells(1, Columns.Count).End(xlToLeft).Column
lastCol = ConvertToLetter(lastColNum)

'Clear content
'Change the value of i
Sheeti.Select
Sheeti.Cells.ClearContens

'Copy and paste
'Change the value of i and j and range
Sheeti.Select    'Sheets("i").Select
Columns("E:E").Select
Selection.Copy
Sheetj.Select    'Sheets("j").Select
Columns("A:A").Select
ActiveSheet.Paste


‘Delete blank rows
Dim lastRow As Integer
lastRow = Cells(Rows.Count, 1).End(xlUp).Row
Dim i As Integer
i = 1
Do Unitl i = lastRow
   If IsEmpty(Range("A" & i).Value) = True Then
      Rows(i & ":" & i).Select
      Selection.Delete Shift:=xlUp
      i = i - 1
      lastRow = lastRow - 1
   End If
   i = i + 1
Loop



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




End Sub


'Function: ConvertToLetter
Function ConvertToLetter(iCol) As String
   Dim iAlpha As Integer
   Dim iRemainder As Integer
   iAlpha = Int(iCol / 27)
   iRemainder = iCol - (iAlpha * 26)
   If iAlpha > 0 Then
      ConvertToLetter = Chr(iAlpha + 64)
   End If
   If iRemainder > 0 Then
      ConvertToLetter = ConvertToLetter & Chr(iRemainder + 64)
   End If
End Function
