Sub customerWeeklyFile()

'1. Copy from full customer file into Sheet1
'2. Find the first "*5/*/2016"(起始日期) in column Y and delete all the above rows in order to avoid overflow
'3. Change the date in the 8 loops
        
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
