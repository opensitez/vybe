' vybe-test: vb/vb_spec_error_handling_resources/error_spec_on_error_resume_next_allows_following_assignment
' origin: languages/vb/tests/vb/test_vb_spec_error_handling_resources.rs

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

Module M
    Sub Main()
        Dim value As Integer = 1
        On Error Resume Next
        Err.Raise(5)
        value = 9
        __Check(CStr(value), "9")
    End Sub
End Module
