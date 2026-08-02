' vybe-test: vb/vb_biginteger_operations/test_vb_bigint_bitwise_and_or_xor
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
        Dim b1 As BigInteger = &HFF00
        Dim b2 As BigInteger = &H00FF
        Dim andRes As BigInteger = b1 And b2
        Dim orRes As BigInteger = b1 Or b2
        __Check(CStr(andRes.ToString()), "0")
        __Check(CStr(orRes.ToString()), "65535")
    End Sub
End Module
