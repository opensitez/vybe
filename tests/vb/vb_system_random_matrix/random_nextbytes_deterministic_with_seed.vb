' vybe-test: vb/vb_system_random_matrix/random_nextbytes_deterministic_with_seed
' origin: languages/vb/tests/vb/test_vb_system_random_matrix.rs

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

Imports System

Module M
    Sub Main()
        Dim a As New Random(99)
        Dim b As New Random(99)
        Dim left(2) As Byte
        Dim right(2) As Byte
        a.NextBytes(left)
        b.NextBytes(right)
        __Check(CStr(left(0) = right(0)), "True")
        __Check(CStr(left(1) = right(1)), "True")
        __Check(CStr(left(2) = right(2)), "True")
    End Sub
End Module
