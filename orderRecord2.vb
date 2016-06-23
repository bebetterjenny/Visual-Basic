Sub order()
    
    Sheets("Shipment").Select
    'MsgBox (Sheet5.Range("B" & 3).Value)
    Dim lastRow1 As Integer
    lastRow1 = Cells(Rows.Count, 1).End(xlUp).Row
    Dim shipmentNo As Variant
    Dim shipmentDate As Variant
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim l As Integer
    
    'Get shipment#
    shipmentNo = InputBox("Please enter the shipment#: ")
    shipmentDate = InputBox("Shipment date: (eg. 6/3/16AIR)")
    
    'MsgBox (shipmentNo)
    'i for Shipment lines
    For i = 2 To 4
    
    MsgBox (i)
    
        For j = 1 To 4
            
            Worksheets(j).Select
            'MsgBox (j)
            Dim lastRow2 As Integer
            lastRow2 = Cells(Rows.Count, 1).End(xlUp).Row
            For k = 2 To lastRow2
                Rows(k).Select
                If Worksheets(j).Range("D" & k).Interior.ColorIndex <> 3 And Worksheets(j).Range("D" & k) = Sheets("Shipment").Range("B" & i) And Worksheets(j).Range("C" & k) = Sheets("Shipment").Range("A" & i) Then
                
                    If Sheets("Shipment").Range("C" & i).Value = "LI" Then
                        
                        If Worksheets(j).Range("L" & k) = 0 Then
                            Worksheets(j).Range("I" & k).Value = shipmentDate
                            Worksheets(j).Range("J" & k).Value = shipmentNo
                            Sheets("Shipment").Select
                            Range("I" & i).Select
                            Selection.Copy
                            Worksheets(j).Select
                            Range("M" & k).Select
                            ActiveSheet.Paste
                        Else
                            l = 1
                            Do Until (Not (Worksheets(j).Range("D" & k + l) = ""))
                            l = l + 1
                            Loop
                            
                            MsgBox ("insert   " & i)
                            
                            Worksheets(j).Select
                            Rows(k + l).Select
                            Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                            Worksheets(j).Range("F" & k + l + 1) = Worksheets(j).Range("K" & k + l)
                            
                            Worksheets(j).Range("I" & k + l).Value = shipmentDate
                            Worksheets(j).Range("J" & k + l).Value = shipmentNo
                            Sheets("Shipment").Select
                            Range("I" & i).Select
                            Selection.Copy
                            Worksheets(j).Select
                            Range("M" & k + l).Select
                            ActiveSheet.Paste
                            
                            
                            Worksheets(j).Range("L" & k + l).Select
                            ActiveCell.FormulaR1C1 = "=SUM(RC[1]:RC[8])"
                            Worksheets(j).Range("K" & k + l).Select
                            ActiveCell.FormulaR1C1 = "=R[-1]C-RC[1]"
    
                            
                            'mark in red
                            If Worksheets(j).Range("K" & k + l).Value = 0 Then
                                Rows(k & ":" & k + l).Select
                                With Selection.Interior
                                    .PatternColorIndex = xlAutomatic
                                    .Color = 255
                                    .TintAndShade = 0
                                    .PatternTintAndShade = 0
                                End With
                            
                            End If
                            
                            
                            k = k + l
                        
                        End If
                        'MsgBox ("LI")

                        
                    ElseIf Sheets("Shipment").Range("C" & i).Value = "M28" Then
                        Worksheets(j).Range("I" & k).Value = shipmentDate
                        Worksheets(j).Range("J" & k).Value = shipmentNo
                        Sheets("Shipment").Select
                        Range("I" & i).Select
                        Selection.Copy
                        Worksheets(j).Select
                        Range("N" & k).Select
                        ActiveSheet.Paste
                        
                        If Worksheets(j).Range("K" & k).Value = 0 Then
                        Rows(k).Select
                            With Selection.Interior
                                .PatternColorIndex = xlAutomatic
                                .Color = 255
                                .TintAndShade = 0
                                .PatternTintAndShade = 0
                            End With
                        End If
                        'MsgBox ("branch")
                        
                        
                        
                    ElseIf Sheets("Shipment").Range("C" & i).Value = "27ST" Then
                        Worksheets(j).Range("I" & k).Value = shipmentDate
                        Worksheets(j).Range("J" & k).Value = shipmentNo
                        Sheets("Shipment").Select
                        Range("I" & i).Select
                        Selection.Copy
                        Worksheets(j).Select
                        Range("O" & k).Select
                        ActiveSheet.Paste
                        
                        If Worksheets(j).Range("K" & k).Value = 0 Then
                        Rows(k).Select
                            With Selection.Interior
                                .PatternColorIndex = xlAutomatic
                                .Color = 255
                                .TintAndShade = 0
                                .PatternTintAndShade = 0
                            End With
                        End If
                        'MsgBox ("branch")
                        
                        
                        
                    ElseIf Sheets("Shipment").Range("C" & i).Value = "L396" Then
                        Worksheets(j).Range("I" & k).Value = shipmentDate
                        Worksheets(j).Range("J" & k).Value = shipmentNo
                        Sheets("Shipment").Select
                        Range("I" & i).Select
                        Selection.Copy
                        Worksheets(j).Select
                        Range("P" & k).Select
                        ActiveSheet.Paste
                        
                        If Worksheets(j).Range("K" & k).Value = 0 Then
                        Rows(k).Select
                            With Selection.Interior
                                .PatternColorIndex = xlAutomatic
                                .Color = 255
                                .TintAndShade = 0
                                .PatternTintAndShade = 0
                            End With
                        End If
                        'MsgBox ("branch")
                        
                        
                        
                    ElseIf Sheets("Shipment").Range("C" & i).Value = "L448" Then
                        Worksheets(j).Range("I" & k).Value = shipmentDate
                        Worksheets(j).Range("J" & k).Value = shipmentNo
                        Sheets("Shipment").Select
                        Range("I" & i).Select
                        Selection.Copy
                        Worksheets(j).Select
                        Range("Q" & k).Select
                        ActiveSheet.Paste
                        
                        If Worksheets(j).Range("K" & k).Value = 0 Then
                        Rows(k).Select
                            With Selection.Interior
                                .PatternColorIndex = xlAutomatic
                                .Color = 255
                                .TintAndShade = 0
                                .PatternTintAndShade = 0
                            End With
                        End If
                        'MsgBox ("branch")
                        
                        
                        
                    ElseIf Sheets("Shipment").Range("C" & i).Value = "DC" Then
                        Worksheets(j).Range("I" & k).Value = shipmentDate
                        Worksheets(j).Range("J" & k).Value = shipmentNo
                        Sheets("Shipment").Select
                        Range("I" & i).Select
                        Selection.Copy
                        Worksheets(j).Select
                        Range("R" & k).Select
                        ActiveSheet.Paste
                                            
                        If Worksheets(j).Range("K" & k).Value = 0 Then
                        Rows(k).Select
                            With Selection.Interior
                                .PatternColorIndex = xlAutomatic
                                .Color = 255
                                .TintAndShade = 0
                                .PatternTintAndShade = 0
                            End With
                        End If
                        'MsgBox ("branch")
                        
                        
                        
                        
                    ElseIf Sheets("Shipment").Range("C" & i).Value = "FL" Then
                        Worksheets(j).Range("I" & k).Value = shipmentDate
                        Worksheets(j).Range("J" & k).Value = shipmentNo
                        Sheets("Shipment").Select
                        Range("I" & i).Select
                        Selection.Copy
                        Worksheets(j).Select
                        Range("S" & k).Select
                        ActiveSheet.Paste
                        
                        If Worksheets(j).Range("K" & k).Value = 0 Then
                        Rows(k).Select
                            With Selection.Interior
                                .PatternColorIndex = xlAutomatic
                                .Color = 255
                                .TintAndShade = 0
                                .PatternTintAndShade = 0
                            End With
                        End If
                        'MsgBox ("branch")
                        
                        
                    ElseIf Sheets("Shipment").Range("C" & i).Value = "TX" Then
                        Worksheets(j).Range("I" & k).Value = shipmentDate
                        Worksheets(j).Range("J" & k).Value = shipmentNo
                        Sheets("Shipment").Select
                        Range("I" & i).Select
                        Selection.Copy
                        Worksheets(j).Select
                        Range("T" & k).Select
                        ActiveSheet.Paste
                        
                        If Worksheets(j).Range("K" & k).Value = 0 Then
                        Rows(k).Select
                            With Selection.Interior
                                .PatternColorIndex = xlAutomatic
                                .Color = 255
                                .TintAndShade = 0
                                .PatternTintAndShade = 0
                            End With
                        End If
                        'MsgBox ("branch")
                    
                    End If
                 
                End If
                
                
                
            Next k
            
        Next j
    Next i
    
End Sub
