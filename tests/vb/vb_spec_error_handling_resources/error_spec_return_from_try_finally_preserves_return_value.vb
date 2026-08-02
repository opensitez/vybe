' vybe-test: vb/vb_spec_error_handling_resources/error_spec_return_from_try_finally_preserves_return_value
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
    Function Work() As Integer
        Try
            Return 5
        Finally
            __Check(CStr("finally"), "finally")
        End Try
    End Function
    Sub Main()
        __Check(CStr(Work()), "5")
    End Sub
End Module
