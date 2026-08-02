' vybe-test: vb/vb_iif_ternary_evaluation/test_vb_if_binary_coalesce_chained
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
        Dim v1 As String = Nothing
        Dim v2 As String = Nothing
        Dim v3 As String = "Final"
        Dim res = If(v1, If(v2, v3))
        __Check(CStr(res), "Final")
    End Sub
End Module
