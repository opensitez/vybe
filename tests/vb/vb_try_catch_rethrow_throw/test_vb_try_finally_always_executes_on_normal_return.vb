' vybe-test: vb/vb_try_catch_rethrow_throw/test_vb_try_finally_always_executes_on_normal_return
' origin: languages/vb/tests/vb/test_vb_try_catch_rethrow_throw.rs

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

Module Program
    Private Function Compute() As String
        Try
            Return "Result"
        Finally
            __Check(CStr("Cleanup in Finally"), "Cleanup in Finally")
        End Try
    End Function

    Sub Main()
        Dim res = Compute()
        __Check(CStr(res), "Result")
    End Sub
End Module
