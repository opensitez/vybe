' vybe-test: vb/vb_big_integer_pow_mod_pow/test_vb_big_integer_mod_pow_crypto_calculation
' origin: languages/vb/tests/vb/test_vb_big_integer_pow_mod_pow.rs

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
        Dim value As New BigInteger(10)
        Dim exponent As New BigInteger(3)
        Dim modulus As New BigInteger(7)
        ' (10^3) mod 7 = 1000 mod 7 = 6
        Dim res = BigInteger.ModPow(value, exponent, modulus)
        __Check(CStr(res.ToString()), "6")
    End Sub
End Module
