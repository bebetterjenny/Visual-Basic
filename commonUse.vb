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