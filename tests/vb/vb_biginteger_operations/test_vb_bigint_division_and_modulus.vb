' vybe-test: vb/vb_biginteger_operations/test_vb_bigint_division_and_modulus
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
        Dim b1 As BigInteger = 1000
        Dim b2 As BigInteger = 300
        Dim div As BigInteger = b1 \ b2
        Dim remVal As BigInteger = b1 Mod b2
        __Check(CStr(div.ToString()), "3")
        __Check(CStr(remVal.ToString()), "100")
    End Sub
End Module
