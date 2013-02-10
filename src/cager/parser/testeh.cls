Public Sub XX()
10    On Error Goto Handler
20       	Thing1
30       	Thing2
40       	Thing3
50    Exit Sub
60    Handler:
70          Err.Raise Err.Number, Err.Source, "XX: " & Err.Description & "(" & Erl & ")"
End Sub

