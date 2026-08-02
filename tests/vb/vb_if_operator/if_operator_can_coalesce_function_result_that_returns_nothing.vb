' vybe-test: vb/vb_if_operator/if_operator_can_coalesce_function_result_that_returns_nothing
' origin: languages/vb/tests/vb/test_vb_if_operator.rs

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
    Function MaybeText(flag As Boolean) As String
        If flag Then
            Return "present"
        End If
        Return Nothing
    End Function

    Sub Main()
        __Check(CStr(If(MaybeText(False), "missing")), "missing")
    End Sub
End Module
