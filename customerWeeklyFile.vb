Sub customerWeeklyFile()

'1. 从原始customer weekly file中copy到Sheet1，然后copy全部paste value（避免日期格式转换成数字）
'2. Y列中找到第一个"*5/*/2016"(起始日期)删掉之前所有行（减小无效数据，避免overflow）
'3. 按顺序更改代码中八个日期为本周范围
        
Sheet2.Select
Sheet2.Cells.ClearContents

Sheet1.Select
Columns("O:O").Select
Selection.Copy
Sheet2.Select
Columns("A:A").Select
ActiveSheet.Paste

Sheet1.Select
Columns("Y:Y").Select
Selection.Copy
Sheet2.Select
Columns("B:B").Select
ActiveSheet.Paste

Sheet1.Select
Columns("Z:Z").Select
Selection.Copy
Sheet2.Select
Columns("C:C").Select
ActiveSheet.Paste

Dim lastRow As Integer
lastRow = Cells(Rows.Count, 1).End(xlUp).Row

Dim copyRowNumber As Integer
copyRowNumber = 2

Dim i As Integer



For i = 2 To lastRow
    If InStr(Sheet2.Range("B" & i), "5/12/2016") Then
        Sheet2.Select
        Rows(i & ":" & i).Select
        Selection.Copy
        Sheet3.Select
        Rows(copyRowNumber & ":" & copyRowNumber).Select
        ActiveSheet.Paste
        Application.CutCopyMode = False
        copyRowNumber = copyRowNumber + 1
    End If
Next

For i = 2 To lastRow
    If InStr(Sheet2.Range("B" & i), "5/13/2016") Then
        Sheet2.Select
        Rows(i & ":" & i).Select
        Selection.Copy
        Sheet3.Select
        Rows(copyRowNumber & ":" & copyRowNumber).Select
        ActiveSheet.Paste
        Application.CutCopyMode = False
        copyRowNumber = copyRowNumber + 1
    End If
Next

For i = 2 To lastRow
    If InStr(Sheet2.Range("B" & i), "5/14/2016") Then
        Sheet2.Select
        Rows(i & ":" & i).Select
        Selection.Copy
        Sheet3.Select
        Rows(copyRowNumber & ":" & copyRowNumber).Select
        ActiveSheet.Paste
        Application.CutCopyMode = False
        copyRowNumber = copyRowNumber + 1
    End If
Next

For i = 2 To lastRow
    If InStr(Sheet2.Range("B" & i), "5/15/2016") Then
        Sheet2.Select
        Rows(i & ":" & i).Select
        Selection.Copy
        Sheet3.Select
        Rows(copyRowNumber & ":" & copyRowNumber).Select
        ActiveSheet.Paste
        Application.CutCopyMode = False
        copyRowNumber = copyRowNumber + 1
    End If
Next

For i = 2 To lastRow
    If InStr(Sheet2.Range("B" & i), "5/16/2016") Then
        Sheet2.Select
        Rows(i & ":" & i).Select
        Selection.Copy
        Sheet3.Select
        Rows(copyRowNumber & ":" & copyRowNumber).Select
        ActiveSheet.Paste
        Application.CutCopyMode = False
        copyRowNumber = copyRowNumber + 1
    End If
Next

For i = 2 To lastRow
    If InStr(Sheet2.Range("B" & i), "5/17/2016") Then
        Sheet2.Select
        Rows(i & ":" & i).Select
        Selection.Copy
        Sheet3.Select
        Rows(copyRowNumber & ":" & copyRowNumber).Select
        ActiveSheet.Paste
        Application.CutCopyMode = False
        copyRowNumber = copyRowNumber + 1
    End If
Next

For i = 2 To lastRow
    If InStr(Sheet2.Range("B" & i), "5/18/2016") Then
        Sheet2.Select
        Rows(i & ":" & i).Select
        Selection.Copy
        Sheet3.Select
        Rows(copyRowNumber & ":" & copyRowNumber).Select
        ActiveSheet.Paste
        Application.CutCopyMode = False
        copyRowNumber = copyRowNumber + 1
    End If
Next

For i = 2 To lastRow
    If InStr(Sheet2.Range("B" & i), "5/19/2016") Then
        Sheet2.Select
        Rows(i & ":" & i).Select
        Selection.Copy
        Sheet3.Select
        Rows(copyRowNumber & ":" & copyRowNumber).Select
        ActiveSheet.Paste
        Application.CutCopyMode = False
        copyRowNumber = copyRowNumber + 1
    End If
Next

    Range("O1").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Sheet3").Select
    Range("A1").Select
    ActiveSheet.Paste
    Sheets("Sheet1").Select
    Range("Y1").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Sheet3").Select
    Range("B1").Select
    ActiveSheet.Paste
    Sheets("Sheet1").Select
    Range("Z1").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Sheet3").Select
    Range("C1").Select
    ActiveSheet.Paste
    
    End Sub
