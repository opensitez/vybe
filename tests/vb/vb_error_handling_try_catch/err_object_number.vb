' vybe-test: vb/vb_error_handling_try_catch/err_object_number
' origin: languages/vb/tests/vb/test_vb_error_handling_try_catch.rs

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
Dim x = 1 \ 0
__Check(CStr(Err.Number <> 0), "True")
End Sub
End Module
