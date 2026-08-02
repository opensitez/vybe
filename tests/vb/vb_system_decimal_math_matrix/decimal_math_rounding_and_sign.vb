' vybe-test: vb/vb_system_decimal_math_matrix/decimal_math_rounding_and_sign
' origin: languages/vb/tests/vb/test_vb_system_decimal_math_matrix.rs

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
        Dim a As Decimal = CDec(10)
        Dim b As Decimal = CDec(3)

        __Check(CStr(Math.Round(a / b, 2)), "3.33")
        __Check(CStr(Math.Sign(-a)), "-1")
        __Check(CStr(Math.Sign(0D)), "0")
        __Check(CStr(Math.Sign(b)), "1")
    End Sub
End Module
