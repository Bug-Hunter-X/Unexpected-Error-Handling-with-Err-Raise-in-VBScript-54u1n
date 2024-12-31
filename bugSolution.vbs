Function MyFunc(param)
  On Error GoTo ErrHandler
  If IsEmpty(param) Then
    Err.Raise vbError + 1, , "Parameter cannot be empty"
  End If
  ' ... rest of the function
  Exit Function
ErrHandler:
  MsgBox "An error occurred: " & Err.Number & " - " & Err.Description, vbCritical
  Err.Clear
End Function