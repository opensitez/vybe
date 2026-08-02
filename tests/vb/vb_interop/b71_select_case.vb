' vybe-test: vb/vb_interop/b71_select_case
' origin: languages/vb/tests/vb/vb_interop_test.rs

' Vybe test harness — Visual Basic.
'
' Real VB source alongside harness/go/check.go and harness/js/check.js, the way
' test262's assert.js is JavaScript.
'
' A test's verdict is its EXIT CODE. __Check prints its diagnostic BEFORE
' throwing: an uncaught exception surfaces as `RuntimeError: [object]`, which
' says nothing at all.

Module VybeCheck
    Sub __Check(got As String, want As String)
        If got <> want Then
            Console.WriteLine("FAIL: want [" & want & "] got [" & got & "]")
            Throw New Exception("assertion failed")
        End If
    End Sub
End Module

Function DayName(d As Integer) As String
    Select Case d
        Case 1
            Return "Monday"
        Case 2
            Return "Tuesday"
        Case 3
            Return "Wednesday"
        Case Else
            Return "Other"
    End Select
End Function
__Check(CStr(DayName(1)), "Monday")
__Check(CStr(DayName(3)), "Wednesday")
__Check(CStr(DayName(7)), "Other")
