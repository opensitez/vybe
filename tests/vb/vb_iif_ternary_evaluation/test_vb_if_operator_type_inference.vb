' vybe-test: vb/vb_iif_ternary_evaluation/test_vb_if_operator_type_inference
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
        ' If operator infers exact common type (Integer)
        Dim num = If(1 > 0, 100, 200)
        __Check(CStr(num + 50), "150")
    End Sub
End Module
