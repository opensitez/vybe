' vybe-test: vb/vb_iif_ternary_evaluation/test_vb_iif_legacy_boolean_coercion_condition
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

Imports Microsoft.VisualBasic

Module Program
    Sub Main()
        ' IIf accepts numeric condition: non-zero is True, 0 is False
        Dim res1 = IIf(1, "TruePart", "FalsePart")
        Dim res2 = IIf(0, "TruePart", "FalsePart")
        __Check(CStr(res1 & "|" & res2), "TruePart|FalsePart")
    End Sub
End Module
