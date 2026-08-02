' vybe-test: vb/vb_linq_all_any_predicates/test_vb_linq_any_short_circuits
' origin: languages/vb/tests/vb/test_vb_linq_all_any_predicates.rs

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

Imports System.Linq

Module Program
    Sub Main()
        Dim count As Integer = 0
        Dim numbers = {10, -5, 20, 30}
        Dim hasNeg = numbers.Any(Function(n)
            count += 1
            Return n < 0
        End Function)
        __Check(CStr(hasNeg & "|count=" & count), "True|count=2")
    End Sub
End Module
