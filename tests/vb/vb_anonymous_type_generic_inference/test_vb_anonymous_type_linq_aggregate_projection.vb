' vybe-test: vb/vb_anonymous_type_generic_inference/test_vb_anonymous_type_linq_aggregate_projection
' origin: languages/vb/tests/vb/test_vb_anonymous_type_generic_inference.rs

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
        Dim numbers = {10, 20, 30}
        Dim stats = numbers.Aggregate(
            New With {.Sum = 0, .Count = 0},
            Function(acc, n) New With {.Sum = acc.Sum + n, .Count = acc.Count + 1}
        )
        __Check(CStr("Sum=" & stats.Sum & "|Count=" & stats.Count), "Sum=60|Count=3")
    End Sub
End Module
