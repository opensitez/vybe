' vybe-test: vb/vb_iif_ternary_evaluation/test_vb_if_operator_short_circuit_evaluation
' origin: languages/vb/tests/vb/test_vb_iif_ternary_evaluation.rs

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
    Private Function SideEffect(msg As String, val As Integer) As Integer
        __Check(CStr("Effect:" & msg), "Effect:TrueBranch")
        Return val
    End Function

    Sub Main()
        ' If ternary operator short-circuits (only true branch evaluated)!
        Dim res = If(True, SideEffect("TrueBranch", 10), SideEffect("FalseBranch", 20))
        __Check(CStr(res), "10")
    End Sub
End Module
