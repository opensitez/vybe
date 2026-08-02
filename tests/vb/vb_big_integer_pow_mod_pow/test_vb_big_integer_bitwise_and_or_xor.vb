' vybe-test: vb/vb_big_integer_pow_mod_pow/test_vb_big_integer_bitwise_and_or_xor
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
        Dim a As New BigInteger(12) ' 1100
        Dim b As New BigInteger(10) ' 1010
        Dim andVal = a And b ' 1000 (8)
        Dim orVal = a Or b   ' 1110 (14)
        Dim xorVal = a Xor b ' 0110 (6)
        __Check(CStr(andVal.ToString() & "|" & orVal.ToString() & "|" & xorVal.ToString()), "8|14|6")
    End Sub
End Module
