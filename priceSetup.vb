Sub priceSetup()

'Make sure to keep the column order in the oringinal sheet as the template below:
'A       B          C        D         E         F         G       H        I        J
'Item    0-INDIV    1-Ret    2_R6pc    3-dist    4-whol    5-VIP   Qty 1    Qty 7    Qty 121
'Split the full product in several excel files
'Do not put over 2800 rows in each oringinal sheet
'Always keep attributes in the top row


Dim lastRow As Integer
lastRow = Cells(Rows.Count, 1).End(xlUp).Row
Dim i As Integer


'Add a new sheet
Sheets.Add(After:=Sheets(Sheets.Count)).Name = "newPrice"
Range("A1").Value = "Item#"
Range("B1").Value = "Price"
Range("C1").Value = "Price Level"
Range("D1").Value = "Quantity"

'Same product starts from 2nd row
For i = 2 To lastRow
Dim j As Integer

    '3 different price levels
    For j = 1 To 3
        Dim k As Integer
        Dim pColNum As Integer
        pColNum = 1
        Dim pCol As String
        pCol = ConvertToLetter(pColNum)
        Dim r As Integer
        Dim s As Integer
        s = 21 * i + 7 * j - 47
        
        
        'A column
        'Locate Item#
        For k = 0 To 6
            r = 21 * i + 7 * j + k - 47
            Sheets("Sheet1").Select
            Range(pCol & i).Select
            Selection.Copy
            Sheets("newPrice").Select
            Range(pCol & r).Select
            ActiveSheet.Paste
        Next
        
        
        'B column
        'Setup price in the same level
        For k = 0 To 6
            r = 21 * i + 7 * j + k - 47
            pColNum = 2 + k
            pCol = ConvertToLetter(pColNum)
            'MsgBox (r)

            'Setup new price
            If k = 6 Then
                pColNum = pColNum + j - 1
                pCol = ConvertToLetter(pColNum)
                'MsgBox (pCol)
                Sheets("Sheet1").Select
                Range(pCol & i).Select
                Selection.Copy
                Sheets("newPrice").Select
                Range("B" & r).Select
                ActiveSheet.Paste
            Else
                'Locate 6 unchanged price in the same level
                'MsgBox (pCol)
                Sheets("Sheet1").Select
                Range(pCol & i).Select
                Selection.Copy
                Sheets("newPrice").Select
                Range("B" & r).Select
                ActiveSheet.Paste
            End If
        Next
        
        
        'C column
        'Get attributes value in C column
        Sheets("newPrice").Select
        Range("C" & s).Value = "0-INDIV"
        Range("C" & (s + 1)).Value = "1-Ret"
        Range("C" & (s + 2)).Value = "2_R6pc"
        Range("C" & (s + 3)).Value = "3-dist"
        Range("C" & (s + 4)).Value = "4-whol"
        Range("C" & (s + 5)).Value = "5-VIP"
        Range("C" & (s + 6)).Value = "6-New"
        
        'D column
        'Set quantity threshold
        Sheets("newPrice").Select
        If j = 1 Then
            For k = 0 To 6
                Range("D" & (s + k)).Value = 1
            Next
        ElseIf j = 2 Then
            For k = 0 To 6
                Range("D" & (s + k)).Value = 7
            Next
        Else
            For k = 0 To 6
                Range("D" & (s + k)).Value = 121
            Next
        End If
        
    Next
    
Next

End Sub

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
