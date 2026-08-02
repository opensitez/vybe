' vybe-test: vb/vb_array_true_for_all_exists/test_vb_array_trueforall_short_circuits
' origin: languages/vb/tests/vb/test_vb_array_true_for_all_exists.rs

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

Module Program
    Sub Main()
        Dim calls As Integer = 0
        Dim numbers As Integer() = {10, -5, 20, 30}
        ' Should fail at second item (-5) and short-circuit
        Dim allPos As Boolean = Array.TrueForAll(numbers, Function(n)
            calls += 1
            Return n > 0
        End Function)
        __Check(CStr(allPos & "|calls=" & calls), "False|calls=2")
    End Sub
End Module
