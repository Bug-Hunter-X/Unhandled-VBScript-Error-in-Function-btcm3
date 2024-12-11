Function MyFunction(param1)
  If IsEmpty(param1) Then
    Err.Raise 5, , "Parameter cannot be empty"
  End If
End Function