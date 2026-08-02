' vybe-test: vb/vb_biginteger_operations/test_vb_bigint_parse_addition
' origin: languages/vb/tests/vb/test_vb_biginteger_operations.rs

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

Imports System.Numerics

Module Program
    Sub Main()
        Dim b1 As BigInteger = BigInteger.Parse("123456789012345678901234567890")
        Dim b2 As BigInteger = BigInteger.Parse("987654321098765432109876543210")
        Dim sum As BigInteger = b1 + b2
        __Check(CStr(sum.ToString()), "1111111110111111111011111111100")
    End Sub
End Module
