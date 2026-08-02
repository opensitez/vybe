' vybe-test: vb/vb_biginteger_operations/test_vb_bigint_pow_function
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
        Dim b As BigInteger = 2
        Dim pow As BigInteger = BigInteger.Pow(b, 64)
        __Check(CStr(pow.ToString()), "18446744073709551616")
    End Sub
End Module
