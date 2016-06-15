Sub instockQty()


'1. make sure there are 4 sheets in the oringinal workbook (unhide if necessary)
'1. 保证Sheet1,Sheet2,Sheet3,Sheet4按顺序放

'2. put full in the Sheet1
'2. 总表copy到Sheet1

'3. put Li in the Sheet2, then run
'3. ChinaTo的LIcopy到Sheet2，然后运行

'4. 检查，删掉不需要的列


'Sort full
Dim lastRow As Integer
lastRow = Sheet1.Cells(Rows.Count, 1).End(xlUp).Row
Columns("A:AS").Select
ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Clear
ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Add Key:=Range("AB2:AB" & lastRow _
    ), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
With ActiveWorkbook.Worksheets("Sheet1").Sort
    .SetRange Range("A1:AS" & lastRow)
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With


'copy from full
Sheet3.Select
Sheet3.Cells.ClearContents

Sheet1.Select
Columns("E:E").Select
Selection.Copy
Sheet3.Select
Columns("A:A").Select
ActiveSheet.Paste

Sheet1.Select
Columns("F:F").Select
Selection.Copy
Sheet3.Select
Columns("B:B").Select
ActiveSheet.Paste

Sheet1.Select
Columns("AB:AB").Select
Selection.Copy
Sheet3.Select
Columns("C:C").Select
ActiveSheet.Paste

Sheet1.Select
Columns("AL:AL").Select
Selection.Copy
Sheet3.Select
Columns("D:D").Select
ActiveSheet.Paste

Sheet1.Select
Columns("F:F").Select
Selection.Copy
Sheet3.Select
Columns("G:G").Select
ActiveSheet.Paste
Sheet3.Cells(1, 7).Value = "sku"


'remove "-W*C" in the sku
Columns("G:G").Select
Selection.Replace What:="-W*C", Replacement:="", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    Application.CutCopyMode = False
    
    
'vlookup from LI-inventory
Range("E1").Select
ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[2],Sheet2!C[-4]:C[52],57,0)"
Dim autoFillRange As String
Let autoFillRange = "E" & "1" & ":" & "E" & lastRow
Range("E1").Select
Selection.AutoFill Destination:=Range(autoFillRange)

'replace #N/A to 0
Sheet3.Cells(1, 5).Value = "qty"
Columns("E:E").Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
Application.CutCopyMode = False
Columns("E:E").Select
Selection.Replace What:="#N/A", Replacement:="0", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    
    
'caculate is_in_stock
Sheet3.Cells(1, 6).Value = "is_in_stock"
Range("F2").Select
ActiveCell.FormulaR1C1 = "=IF(RC[-1]>70,1,0)"
Let autoFillRange = "F" & "2" & ":" & "F" & lastRow
Range("F2").Select
Selection.AutoFill Destination:=Range(autoFillRange)
Range(autoFillRange).Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
Application.CutCopyMode = False


'delete sku without -W*C
Columns("G:G").Select
Selection.ClearContents


'check sbs and vr-
Dim i As Integer
Dim checkRange1 As String
Dim checkRange2 As String
Dim changRange As String
'Dim lastRow As Integer
lastRow = Sheet3.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastRow
Let checkRange1 = "B" & i
Let checkRange2 = "E" & i
Let changeRange = "F" & i

If (InStr(Sheet3.Range(checkRange1), "SBS") Or InStr(Sheet3.Range(checkRange2), "VR-") Or InStr(Sheet3.Range(checkRange2), "SCP102-")) And Sheet3.Range(checkRange2) > 3 Then

    Sheet3.Range(changeRange).Value = 1
    
End If
Next

Range("A1").Select



'move sku that is both simple and catalog, search
'Dim i As Integer
Dim copyRowNumber As Integer
copyRowNumber = 1


For i = 2 To lastRow
    If (InStr(Sheet3.Range("A" & i), "simple") Or InStr(Sheet3.Range("A" & i), "Simple")) And InStr(Sheet3.Range("D" & i), "Catalog, Search") Then
        Rows(i & ":" & i).Select
        Selection.Copy
        Sheets("Sheet4").Select
        Rows(copyRowNumber & ":" & copyRowNumber).Select
        ActiveSheet.Paste
        Application.CutCopyMode = False
        copyRowNumber = copyRowNumber + 1
        Sheets("Sheet3").Select
        Rows(i & ":" & i).Delete
        i = i - 1
    End If
Next


'cut names and text to columns

Columns("C:C").Select
Selection.Copy
Columns("G:G").Select
ActiveSheet.Paste
Selection.Replace What:="in*1", Replacement:="", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
Selection.Replace What:=" In ", Replacement:="&", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
        
Columns("G:G").Select
Selection.TextToColumns Destination:=Range("G1"), DataType:=xlDelimited, _
    TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
    Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
    :="&", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True


'subtotal

Columns("A:G").Select
Selection.Subtotal GroupBy:=7, Function:=xlMax, TotalList:=Array(6), _
    Replace:=True, PageBreaks:=False, SummaryBelowData:=False
    
Columns("G:J").Delete


'set config is_in_stock

'Dim lastRow As Integer
lastRow = Sheet3.Cells(Rows.Count, 1).End(xlUp).Row

For i = 3 To lastRow
    Sheet3.Range("G" & i).Select
    ActiveCell.FormulaR1C1 = "=R[-1]C[-1]"
Next

Range("G:G").Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    
Rows(lastRow + 1 & ":" & lastRow + 1).Delete

Cells.Select
Selection.RemoveSubtotal

For i = 2 To lastRow
    If Sheet3.Range("A" & i) = "configurable" Then
        Sheet3.Range("F" & i).Select
        ActiveCell.FormulaR1C1 = "=R[0]C[1]"
    End If
Next


'Copy simple and Catalog, Search back into Sheet3

Dim lastRowSheet3 As Integer
lastRowSheet3 = Sheet3.Cells(Rows.Count, 1).End(xlUp).Row
Dim lastRowSheet4 As Integer
lastRowSheet4 = Sheet4.Cells(Rows.Count, 1).End(xlUp).Row

    Sheets("Sheet4").Select
    Rows(1 & ":" & lastRowSheet4).Select
    Selection.Copy
    Sheets("Sheet3").Select
    Rows((lastRowSheet3 + 1) & ":" & (lastRowSheet3 + lastRowSheet4)).Select
    ActiveSheet.Paste


	
'Sort
lastRow = Sheet3.Cells(Rows.Count, 1).End(xlUp).Row
Columns("A:G").Select
ActiveWorkbook.Worksheets("Sheet3").Sort.SortFields.Clear
ActiveWorkbook.Worksheets("Sheet3").Sort.SortFields.Add Key:=Range("B2:B" & lastRow _
    ), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
With ActiveWorkbook.Worksheets("Sheet3").Sort
    .SetRange Range("A1:G" & lastRow)
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With


'Paste value and delete the last column
Columns("G:G").Select
Cells.Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
Columns("G:G").Select
Application.CutCopyMode = False
Selection.Delete Shift:=xlToLeft

Msgbox("1. 检查sbs，vr-，SCP102-;检查母体计算是否正确" & vbNewLine & "2. 删除A，C，D列")


End Sub
