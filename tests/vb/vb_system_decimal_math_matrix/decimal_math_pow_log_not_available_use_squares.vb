' vybe-test: vb/vb_system_decimal_math_matrix/decimal_math_pow_log_not_available_use_squares
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
        Dim x As Decimal = CDec(2)
        Dim squared As Decimal = x * x
        Dim cubed As Decimal = x * x * x

        __Check(CStr(squared), "4")
        __Check(CStr(cubed), "8")
    End Sub
End Module
