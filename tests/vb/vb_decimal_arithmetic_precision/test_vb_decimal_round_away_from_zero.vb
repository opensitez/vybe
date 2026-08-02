' vybe-test: vb/vb_decimal_arithmetic_precision/test_vb_decimal_round_away_from_zero
' origin: languages/vb/tests/vb/test_vb_decimal_arithmetic_precision.rs

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
        Dim r1 As Decimal = Decimal.Round(2.5D, 0, MidpointRounding.AwayFromZero)
        Dim r2 As Decimal = Decimal.Round(3.5D, 0, MidpointRounding.AwayFromZero)
        __Check(CStr(r1), "3")
        __Check(CStr(r2), "4")
    End Sub
End Module
