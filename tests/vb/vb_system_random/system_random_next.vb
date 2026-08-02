' vybe-test: vb/vb_system_random/system_random_next
' origin: languages/vb/tests/vb/test_vb_system_random.rs

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
        Dim r As New Random(42) ' Seeded for determinism
        Dim val1 = r.Next(1, 100)
        Dim val2 = r.Next(1, 100)
        
        __Check(CStr(val1 >= 1 AndAlso val1 < 100), "True")
        __Check(CStr(val2 >= 1 AndAlso val2 < 100), "True")
        
        Dim r2 As New Random(42)
        Dim val3 = r2.Next(1, 100)
        __Check(CStr(val1 = val3), "True") ' Deterministic with same seed
    End Sub
End Module
