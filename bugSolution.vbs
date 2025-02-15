Function CreateObject(progID)
  On Error Resume Next
  Dim obj
  Set obj = GetObject(, progID)
  If Err.Number <> 0 Then
    Err.Clear
    Set obj = CreateObject(progID)
    If Err.Number <> 0 Then
      MsgBox "Error creating object: " & Err.Description, vbCritical
      Exit Function
    End If
  End If
  Set CreateObject = obj
End Function

' Example usage:
Dim myObject
Set myObject = CreateObject("MyCOMObject.MyClass")
If IsObject(myObject) Then
  'Use the myObject
Else
  'Handle the error
End If