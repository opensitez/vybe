' vybe-test: vb/vb_spec_error_handling_resources/error_spec_nested_try_outer_catch_handles_unhandled_inner_exception
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
        Try
            Try
                Throw New Exception("inner")
            Finally
                __Check(CStr("inner finally"), "inner finally")
            End Try
        Catch ex As Exception
            __Check(CStr("outer catch"), "outer catch")
        End Try
    End Sub
End Module
