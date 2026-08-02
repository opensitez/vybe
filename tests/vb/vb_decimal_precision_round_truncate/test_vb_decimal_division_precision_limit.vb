' vybe-test: vb/vb_decimal_precision_round_truncate/test_vb_decimal_division_precision_limit
' origin: languages/vb/tests/vb/test_vb_decimal_precision_round_truncate.rs

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
        Dim a As Decimal = 1D
        Dim b As Decimal = 3D
        Dim res = a / b
        __Check(CStr(res.ToString().StartsWith("0.3333333333333333333333333333")), "True")
    End Sub
End Module
