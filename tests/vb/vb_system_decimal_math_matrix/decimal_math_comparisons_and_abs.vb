' vybe-test: vb/vb_system_decimal_math_matrix/decimal_math_comparisons_and_abs
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
        Dim a As Decimal = CDec(-12.34)

        __Check(CStr(Decimal.Compare(a, Decimal.Zero) < 0), "True")
        __Check(CStr(Decimal.MaxValue > a), "True")
        __Check(CStr(Decimal.MinValue < a), "True")
        __Check(CStr(Decimal.Round(Math.Abs(a), 1)), "12.3")
    End Sub
End Module
