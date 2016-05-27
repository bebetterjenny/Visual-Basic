Sub saleReport()

'OMS >> Wildcard Search >> Start: am.csv; pm.csv
'OMS >> Inventory Log Inquiry >> From 前一天日期 To 前一天日期, W# 分店 To 同一分店 >> Start: DC.csv; FL.csv; L396.csv; L448.csv; TX.csv
'拿出Sale Report Tool.xlsm与am.csv; pm.csv; DC.csv; FL.csv; L396.csv; L448.csv; TX.csv放在同一文件夹
'Open all the 8 files

'Copy from csv
'am
Windows("am.CSV").Activate
Cells.Select
Selection.Copy
Windows("Sale Report Tool.xlsm").Activate
Sheets("am").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
'pm
Windows("pm.CSV").Activate
Cells.Select
Application.CutCopyMode = False
Selection.Copy
Windows("Sale Report Tool.xlsm").Activate
Sheets("pm").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
'DC
Windows("DC.CSV").Activate
Cells.Select
Application.CutCopyMode = False
Selection.Copy
Windows("Sale Report Tool.xlsm").Activate
Sheets("DCy").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
'FL
Windows("FL.CSV").Activate
Cells.Select
Application.CutCopyMode = False
Selection.Copy
Windows("Sale Report Tool.xlsm").Activate
Sheets("FLy").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
'L396
Windows("L396.CSV").Activate
Cells.Select
Application.CutCopyMode = False
Selection.Copy
Windows("Sale Report Tool.xlsm").Activate
Sheets("L396y").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
'448
Windows("L448.CSV").Activate
Cells.Select
Application.CutCopyMode = False
Selection.Copy
Windows("Sale Report Tool.xlsm").Activate
Sheets("L448y").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
'TX
Windows("TX.CSV").Activate
Cells.Select
Application.CutCopyMode = False
Selection.Copy
Windows("Sale Report Tool.xlsm").Activate
Sheets("TXy").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    
    
Dim lastRow As Integer
'am
Sheets("am").Select
Range("C:G,I:O,U:U").Select
Selection.Delete Shift:=xlToLef
lastRow = Cells(Rows.Count, 1).End(xlUp).Row
Columns("B:B").Select
ActiveWorkbook.Worksheets("am").Sort.SortFields.Clear
ActiveWorkbook.Worksheets("am").Sort.SortFields.Add Key:=Range("B1"), SortOn _
    :=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
With ActiveWorkbook.Worksheets("am").Sort
    .SetRange Range("A2:H" & lastRow)
    .Header = xlNo
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With
Selection.AutoFilter
ActiveSheet.Range("$B$1:$B$" & lastRow).AutoFilter Field:=1, Criteria1:=Array( _
    "L396", "L448", "TX"), Operator:=xlFilterValues
'pm
Sheets("pm").Select
Range("C:G,I:O,U:U").Select
Selection.Delete Shift:=xlToLef
lastRow = Cells(Rows.Count, 1).End(xlUp).Row
Columns("B:B").Select
ActiveWorkbook.Worksheets("pm").Sort.SortFields.Clear
ActiveWorkbook.Worksheets("pm").Sort.SortFields.Add Key:=Range("B1"), SortOn _
    :=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
With ActiveWorkbook.Worksheets("pm").Sort
    .SetRange Range("A2:H" & lastRow)
    .Header = xlNo
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With
Selection.AutoFilter
ActiveSheet.Range("$B$1:$B$" & lastRow).AutoFilter Field:=1, Criteria1:=Array( _
    "DC", "FL", "LI"), Operator:=xlFilterValues
    
'DCy
Sheets("DCy").Select
lastRow = Cells(Rows.Count, 1).End(xlUp).Row
Columns("D:D").Select
ActiveWorkbook.Worksheets("DCy").Sort.SortFields.Clear
ActiveWorkbook.Worksheets("DCy").Sort.SortFields.Add Key:=Range("D1"), _
    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
With ActiveWorkbook.Worksheets("DCy").Sort
    .SetRange Range("A2:M" & lastRow)
    .Header = xlNo
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With
Selection.AutoFilter
'FLy
Sheets("FLy").Select
lastRow = Cells(Rows.Count, 1).End(xlUp).Row
Columns("D:D").Select
ActiveWorkbook.Worksheets("FLy").Sort.SortFields.Clear
ActiveWorkbook.Worksheets("FLy").Sort.SortFields.Add Key:=Range("D1"), _
    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
With ActiveWorkbook.Worksheets("FLy").Sort
    .SetRange Range("A2:M" & lastRow)
    .Header = xlNo
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With
Selection.AutoFilter
'L396y
Sheets("L396y").Select
lastRow = Cells(Rows.Count, 1).End(xlUp).Row
Columns("D:D").Select
ActiveWorkbook.Worksheets("L396y").Sort.SortFields.Clear
ActiveWorkbook.Worksheets("L396y").Sort.SortFields.Add Key:=Range("D1"), _
    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
With ActiveWorkbook.Worksheets("L396y").Sort
    .SetRange Range("A2:M" & lastRow)
    .Header = xlNo
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With
Selection.AutoFilter
'L448y
Sheets("L448y").Select
lastRow = Cells(Rows.Count, 1).End(xlUp).Row
Columns("D:D").Select
ActiveWorkbook.Worksheets("L448y").Sort.SortFields.Clear
ActiveWorkbook.Worksheets("L448y").Sort.SortFields.Add Key:=Range("D1"), _
    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
With ActiveWorkbook.Worksheets("L448y").Sort
    .SetRange Range("A2:M" & lastRow)
    .Header = xlNo
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With
Selection.AutoFilter
'TXy
Sheets("TXy").Select
lastRow = Cells(Rows.Count, 1).End(xlUp).Row
Columns("D:D").Select
ActiveWorkbook.Worksheets("TXy").Sort.SortFields.Clear
ActiveWorkbook.Worksheets("TXy").Sort.SortFields.Add Key:=Range("D1"), _
    SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
With ActiveWorkbook.Worksheets("TXy").Sort
    .SetRange Range("A2:M" & lastRow)
    .Header = xlNo
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With
Selection.AutoFilter
    
MsgBox ("1. am,pm选中copy paste value》filter》name range." & vbNewLine & "2. 分店选出所需要的copy")


    


End Sub