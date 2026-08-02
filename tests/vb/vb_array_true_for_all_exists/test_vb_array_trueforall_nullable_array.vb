' vybe-test: vb/vb_array_true_for_all_exists/test_vb_array_trueforall_nullable_array
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
        Dim items As Nullable(Of Integer)() = {10, 20, 30}
        Dim allHasValue As Boolean = Array.TrueForAll(items, Function(i) i.HasValue)
        __Check(CStr(allHasValue), "True")
    End Sub
End Module
