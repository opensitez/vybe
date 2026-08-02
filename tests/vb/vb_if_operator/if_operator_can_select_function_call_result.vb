' vybe-test: vb/vb_if_operator/if_operator_can_select_function_call_result
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
    Function AddOne(value As Integer) As Integer
        Return value + 1
    End Function

    Sub Main()
        __Check(CStr(If(True, AddOne(4), AddOne(9))), "5")
    End Sub
End Module
