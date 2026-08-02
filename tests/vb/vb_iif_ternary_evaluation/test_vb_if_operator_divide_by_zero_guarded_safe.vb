' vybe-test: vb/vb_iif_ternary_evaluation/test_vb_if_operator_divide_by_zero_guarded_safe
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
    Sub Main()
        Dim divisor = 0
        ' If ternary short-circuits so false branch 10 \ divisor is NOT evaluated!
        Dim res = If(divisor <> 0, 10 \ divisor, -1)
        __Check(CStr(res), "-1")
    End Sub
End Module
