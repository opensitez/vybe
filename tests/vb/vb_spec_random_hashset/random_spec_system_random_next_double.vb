' vybe-test: vb/vb_spec_random_hashset/random_spec_system_random_next_double
' origin: languages/vb/tests/vb/test_vb_spec_random_hashset.rs

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

Module M
    Sub Main()
        __Check(CStr(Imports System
Module Program
    Sub Main()
        Dim rng As New Random()
        Dim value As Double = rng.NextDouble()
        Console.WriteLine(value >= 0 AndAlso value < 1)
    End Sub
End Module), "True")
    End Sub
End Module
