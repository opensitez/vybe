' vybe-test: vb/vb_array_find_findindex_predicates/test_vb_array_findlastindex_start_index_backwards
' origin: languages/vb/tests/vb/test_vb_array_find_findindex_predicates.rs

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
        Dim numbers As Integer() = {10, 20, 30, 20, 10}
        ' Search backwards from index 2
        Dim idx As Integer = Array.FindLastIndex(numbers, 2, Function(n) n = 20)
        __Check(CStr(idx), "1")
    End Sub
End Module
