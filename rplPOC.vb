Sub CCCCCC()

'A       B             C            D      E                   F
'color   model short   model long   name   short_description   description

Dim lastRow As Integer
lastRow = Cells(Rows.Count, 1).End(xlUp).Row
Dim i As Integer
Dim rplText As String

For i = 2 To lastRow
    If Range("A" & i).Value = "" Then
        rplText = Replace(Range("D" & i).Value, " In CCCCCC", "")
        Range("D" & i).Value = rplText
        rplText = Replace(Range("E" & i).Value, " in CCCCCC", "")
        Range("E" & i).Value = rplText
        rplText = Replace(Range("E" & i).Value, " In CCCCCC", "")
        Range("E" & i).Value = rplText
    Else
        rplText = Replace(Range("D" & i).Value, "CCCCCC", Range("A" & i).Value)
        Range("D" & i).Value = rplText
        rplText = Replace(Range("E" & i).Value, "CCCCCC", LCase(Range("A" & i).Value))
        Range("E" & i).Value = rplText
    End If
Next
        

End Sub
Sub OOOOOO()

Dim lastRow As Integer
lastRow = Cells(Rows.Count, 1).End(xlUp).Row
Dim i As Integer
Dim rplText As String

'model short
For i = 2 To lastRow
    rplText = Replace(Range("D" & i).Value, "OOOOOO", Range("B" & i).Value)
    Range("D" & i).Value = rplText
    rplText = Replace(Range("E" & i).Value, "OOOOOO", Range("B" & i).Value)
    Range("E" & i).Value = rplText
Next

'model long
For i = 2 To lastRow

    If Range("C" & i).Value = "" Then
        rplText = Replace(Range("F" & i).Value, "</li><br/><li>This phone model is also known as OOOOOO", "")
        Range("F" & i).Value = rplText
    ElseIf InStr(Range("C" & i), "iPhone") Then
        rplText = Replace(Range("F" & i).Value, "This phone model is also known as OOOOOO", Range("C" & i).Value)
        Range("F" & i).Value = rplText
    Else
        rplText = Replace(Range("F" & i).Value, "OOOOOO", Range("C" & i).Value)
        Range("F" & i).Value = rplText
    End If
Next


End Sub