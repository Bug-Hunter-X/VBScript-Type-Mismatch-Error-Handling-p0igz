Function MyFunction(param1, param2)
  On Error Resume Next
  'Check for null or empty values
  If IsNull(param1) Or param1 = "" Then
    Err.Raise 13, , "Type mismatch: param1 is null or empty"
  End If
  If IsNull(param2) Or param2 = "" Then
    Err.Raise 13, , "Type mismatch: param2 is null or empty"
  End If
  'Check for numeric values if needed
  If Not IsNumeric(param1) Then
    Err.Raise 13, , "Type mismatch: param1 is not numeric"
  End If
  If Not IsNumeric(param2) Then
    Err.Raise 13, , "Type mismatch: param2 is not numeric"
  End If
  On Error GoTo 0
  'Rest of the code
End Function