' vybe-test: vb/vb_biginteger_operations/test_vb_bigint_greatest_common_divisor
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
        Dim b1 As BigInteger = 54
        Dim b2 As BigInteger = 24
        Dim gcd As BigInteger = BigInteger.GreatestCommonDivisor(b1, b2)
        __Check(CStr(gcd.ToString()), "6")
    End Sub
End Module
