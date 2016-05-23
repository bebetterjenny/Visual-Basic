Sub fileRenamer()

'Original file name: a1.jpg, a2.jpg, a3.jpg, etc.
'Rename: b1.jpg, b2.jpg, b3.jpg, etc.

Dim r As String
Dim i As Integer
i = 1

'Change the path and the file
r = Dir("I:\temp\*.jpg")
MsgBox ("The first file name is: " & r)

Do While r <> ""
o = r
'Change the name
r = "b" & i & ".jpg"
Name "I:\temp\" & o As "I:\temp\" & r
i = i + 1

r = Dir()
'MsgBox (r)
Loop

r = Dir("I:\temp\*.jpg")
MsgBox ("Changed. The first file name is: " & r)

End Sub
