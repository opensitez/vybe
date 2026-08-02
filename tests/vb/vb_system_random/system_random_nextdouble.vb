' vybe-test: vb/vb_system_random/system_random_nextdouble
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
        Dim r As New Random()
        Dim val = r.NextDouble()
        __Check(CStr(val >= 0.0 AndAlso val < 1.0), "True")
    End Sub
End Module
