' vybe-test: vb/vb_iif_ternary_evaluation/test_vb_iif_returns_object_requiring_conversion
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
        ' IIf returns Object type
        Dim objRes = IIf(1 > 0, 100, 200)
        Dim num As Integer = CInt(objRes)
        __Check(CStr(num), "100")
    End Sub
End Module
