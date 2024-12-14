Function MyFunction(param1)
  On Error Resume Next
  'Improved parameter handling
  If IsEmpty(param1) Then
    MsgBox "Parameter is empty", vbInformation
  ElseIf IsNumeric(param1) Then
    ' Handle numeric parameter
  ElseIf VarType(param1) = vbString Then
    ' Handle string parameter
  Else
    ErrNumber = Err.Number
    If ErrNumber <> 0 Then
       MsgBox "Error: " & Err.Description, vbCritical
       Err.Clear 'Clear the error to avoid subsequent problems
    End If
  End If
End Function