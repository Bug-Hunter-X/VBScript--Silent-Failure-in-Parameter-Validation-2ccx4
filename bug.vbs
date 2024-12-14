Function MyFunction(param1)
  If IsEmpty(param1) Then
    ' Handle empty parameter
  ElseIf IsNumeric(param1) Then
    ' Handle numeric parameter
  ElseIf VarType(param1) = vbString Then
    ' Handle string parameter
  Else
    Err.Raise vbObjectError + 5000, "MyFunction", "Invalid parameter type"
  End If
End Function