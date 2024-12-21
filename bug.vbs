Function MyFunction(param1, param2)
  'Some code here that may cause an error
  If param1 = "" Then
    Err.Raise 13, , "Type mismatch"
  End If
  'More code here
End Function