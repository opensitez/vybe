' vybe-test: vb/vb_decimal_arithmetic_precision/test_vb_decimal_get_bits_array
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
        Dim d As Decimal = 10.0D
        Dim bits As Integer() = Decimal.GetBits(d)
        __Check(CStr(bits.Length), "4")
    End Sub
End Module
