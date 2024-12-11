Function MyFunction(param1)
  On Error Resume Next
  If IsEmpty(param1) Then
    Err.Raise 5, , "Parameter cannot be empty"
  End If
  On Error GoTo 0
End Function

Sub Main()
  On Error GoTo ErrorHandler
  Dim result
  result = MyFunction("")
  If Err.Number <> 0 Then
    MsgBox "Error: " & Err.Description
  Else
    MsgBox "Function executed successfully"
  End If
  Exit Sub
ErrorHandler:
  MsgBox "An error occurred: " & Err.Description
End Sub