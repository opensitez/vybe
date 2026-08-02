' vybe-test: vb/vb_error_handling_try_catch/try_catch_throw_no_arg
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
Try
Try
Throw New System.Exception()
Catch
Throw
End Try
Catch
__Check(CStr("Rethrown"), "Rethrown")
End Try
End Sub
End Module
