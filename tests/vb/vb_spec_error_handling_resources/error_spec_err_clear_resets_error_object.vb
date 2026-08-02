' vybe-test: vb/vb_spec_error_handling_resources/error_spec_err_clear_resets_error_object
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
        On Error Resume Next
        Err.Raise(5, , "boom")
        __Check(CStr(Err.Description), "boom")
        Err.Clear()
        __Check(CStr(Err.Number), "0")
    End Sub
End Module
