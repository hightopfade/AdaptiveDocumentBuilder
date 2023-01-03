Sub Module1()

CreateVBAModules (1)
out = RenameVBAModule("Module1", "testing")

End Sub

Function RandomString(Length As Integer)

Dim CharacterBank As Variant
Dim x As Long
Dim str As String

If Length < 1 Then
    MsgBox "Length variable must be greater than 0"
    Exit Function
End If

CharacterBank = Array("a", "b", "c", "d", "e", "f", "g", "h", "i", "j", _
    "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", _
    "y", "z", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", _
    "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z")
  

For x = 1 To Length
    Randomize
    str = str & CharacterBank(Int((UBound(CharacterBank) - LBound(CharacterBank) + 1) * Rnd + LBound(CharacterBank)))
Next x


RandomString = str

End Function

Function CreateVBAModules(itr As Integer)

Set wdDoc = ActiveDocument

cnt = 1
While cnt <= itr
    wdDoc.VBProject.VBComponents.Add (1)
    cnt = cnt + 1
Wend

End Function

Function RenameVBAModule(fromName As String, toName As String)

Set wdDoc = ActiveDocument
Set rename = wdDoc.VBProject.VBComponents.Item(fromName)
rename.Name = toName

End Function