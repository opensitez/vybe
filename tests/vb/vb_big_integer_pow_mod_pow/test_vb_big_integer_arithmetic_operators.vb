' vybe-test: vb/vb_big_integer_pow_mod_pow/test_vb_big_integer_arithmetic_operators
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
        Dim a As BigInteger = 1000000000000000000L
        Dim b As BigInteger = 2000000000000000000L
        Dim sum = a + b
        Dim diff = b - a
        Dim mult = a * 2
        __Check(CStr(sum.ToString() & "|" & diff.ToString() & "|" & mult.ToString()), "3000000000000000000|1000000000000000000|2000000000000000000")
    End Sub
End Module
