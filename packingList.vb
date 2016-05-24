Sub packingList()


'1.Remove branches in the original packinglist
'1.在原始装箱单中去掉分店

'2.Copy and paste value from original packinglist into Sheet1
'2.复制原始装箱单中全部值到Sheet1

'3.Check new items and copy them into Sheet3
'3.把查到的新item复制到Sheet3

'4.After running all the packingList(),fill the Sheet3,and save as xls/xlsx
'4.跑完此程序后在Sheet3中手动填充完整，另存为xls/xlsx



'Sheet2 is the calculated packinglist
'Sheet3
'Copy from packingList
    Sheet2.Select
    Sheet2.Cells.ClearContents
    
    Sheet1.Select
    Range("C:E,G:G,N:N,S:U").Select
    Selection.Copy
    Sheets("Sheet2").Select
    Range("A1").Select
    ActiveSheet.Paste
    Range("A1").Value = "Item"
    Range("B1").Value = "Qty"
    Range("C1").Value = "Qty/Carton"
    Range("D1").Value = "Weight(lb)/Carton"
    Range("E1").Value = "Dimension(mm)/Unit"
    Range("F1").Value = "L/Case"
    Range("G1").Value = "W/Case"
    Range("H1").Value = "H/Case"
    Rows("2:11").Select
    Selection.Delete Shift:=xlUp
    
    Dim lastRow As Integer
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Dim i As Integer
    i = 1
    
    Do Until i = lastRow
        i = i + 1
        If IsEmpty(Range("B" & i).Value) = True Or InStr(1, Range("A" & i), "Z") = 1 Or InStr(Range("A" & i), "SCP102") Or InStr(Range("A" & i), "Carton") > 0 Then
            Rows(i & ":" & i).Select
            Selection.Delete Shift:=xlUp
            i = i - 1
            lastRow = lastRow - 1
            'MsgBox (i)
        End If
    Loop
    
    
    'Sumif and remove duplicate
    Range("P2").Select
    ActiveCell.FormulaR1C1 = "=SUMIF(C[-15],C[-15],C[-14])"
    Range("P2").Select
    Selection.AutoFill Destination:=Range("P2:P" & lastRow)
    Range("P2:P" & lastRow).Select
    Selection.Copy
    Range("B2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("A:O").Select
    Application.CutCopyMode = False
    ActiveSheet.Range("$A$1:$O$" & lastRow).RemoveDuplicates Columns:=1, Header:=xlYes
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    Columns("P:P").Select
    Selection.Delete Shift:=xlToLeft
    
    
    'Remove "mm"
    Range("E2:E" & lastRow).Select
    'Range(Selection, Selection.End(xlDown)).Select
    Selection.Replace What:="mm", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
     'Insert 6 columns
     Range("F:F").Select
     Selection.Resize(, 6).EntireColumn.Insert Shift:=xlRight
     
     
     'Dilemma
    Columns("E:E").Select
    Selection.TextToColumns Destination:=Range("F1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="*", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1)), _
        TrailingMinusNumbers:=True

    Range("F1").Value = "l(mm)/Unit"
    Range("G1").Value = "w(mm)/Unit"
    Range("H1").Value = "h(mm)/Unit"
    Range("I1").Value = "l(inch)/Unit"
    Range("J1").Value = "w(inch)/Unit"
    Range("K1").Value = "h(inch)/Unit"
    
    
    'mm to inch and get the left 3 digits after decimal points
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "=LEFT(0.0393*RC[-3],5)"
    Selection.AutoFill Destination:=Range("I2:K2"), Type:=xlFillDefault
    Range("I2:K2").Select
    Selection.AutoFill Destination:=Range("I2:K" & lastRow)

    
    'Calculate weight
    Range("O1").Value = "Weight(lb)/Unit"
    Range("O2").Select
    ActiveCell.FormulaR1C1 = "=LEFT(RC[-11]/RC[-12],4)"
    Range("O2").Select
    Selection.AutoFill Destination:=Range("O2:O" & lastRow)
    Range("P1").Value = "Weight(oz)/Unit"
    Range("P2").Select
    ActiveCell.FormulaR1C1 = "=LEFT(RC[-1]*16,4)"
    Range("P2").Select
    Selection.AutoFill Destination:=Range("P2:P" & lastRow)

        
    'Sort
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    ActiveWorkbook.Worksheets("Sheet2").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet2").Sort.SortFields.Add Key:=Range("A2:A" & lastRow) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Sheet2").Sort
        .SetRange Range("A1:O" & lastRow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
        
    
    'New items in Sheet3
    Sheet3.Select
    Range("A1").Value = "item"
    Range("B1").Value = "l/Unit"
    Range("C1").Value = "w/Unit"
    Range("D1").Value = "h/Unit"
    Range("E1").Value = "l/Case"
    Range("F1").Value = "w/Case"
    Range("G1").Value = "h/Case"
    Range("H1").Value = "Weight(lb)/Qty"
    Range("I1").Value = "Weight(lb)/Case"
    Range("J1").Value = "oz"
    
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(Sheet3!RC1,Sheet2!C1:C16,9,0)"
    Range("C2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(Sheet3!RC1,Sheet2!C1:C16,10,0)"
    Range("D2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(Sheet3!RC1,Sheet2!C1:C16,11,0)"
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(Sheet3!RC1,Sheet2!C1:C16,12,0)"
    Range("F2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(Sheet3!RC1,Sheet2!C1:C16,13,0)"
    Range("G2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(Sheet3!RC1,Sheet2!C1:C16,14,0)"
    Range("H2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(Sheet3!RC1,Sheet2!C1:C16,3,0)"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(Sheet3!RC1,Sheet2!C1:C16,15,0)"
    Range("J2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(Sheet3!RC1,Sheet2!C1:C16,16,0)"
    
    
    
    
End Sub



