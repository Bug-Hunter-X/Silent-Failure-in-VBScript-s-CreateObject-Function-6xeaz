Function CreateObject(progID) 
Dim obj
Set obj = GetObject( , progID) 
If IsObject(obj) Then
  Set CreateObject = obj
Else
  Set CreateObject = CreateObject(progID)
End If
End Function