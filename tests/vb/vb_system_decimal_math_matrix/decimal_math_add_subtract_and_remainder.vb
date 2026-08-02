' vybe-test: vb/vb_system_decimal_math_matrix/decimal_math_add_subtract_and_remainder
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
        Dim a As Decimal = CDec(12.75)
        Dim b As Decimal = CDec(0.75)

        __Check(CStr(a + b), "13.5")
        __Check(CStr(a - b), "12")
        __Check(CStr(a Mod b), "0")
    End Sub
End Module
