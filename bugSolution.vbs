Function MyFunction(param)
  On Error Resume Next
  If IsEmpty(param) Then
    Err.Raise 9, , "Parameter cannot be empty"
  End If
  On Error GoTo 0
  ' ... rest of the function ...
End Function

Sub Main()
  On Error Resume Next
  Dim result
  result = MyFunction("")
  If Err.Number <> 0 Then
    MsgBox "Error: " & Err.Description
  Else
    MsgBox "Function executed successfully"
  End If
  On Error GoTo 0
End Sub