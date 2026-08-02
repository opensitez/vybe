' vybe-test: vb/vb_array_find_findindex_predicates/test_vb_array_findindex_start_index_and_count_no_match
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
        Dim numbers As Integer() = {2, 4, 6, 8, 10}
        ' Search range [0, 2] = indices 0 and 1 (values 2, 4)
        Dim idx As Integer = Array.FindIndex(numbers, 0, 2, Function(n) n > 5)
        __Check(CStr(idx), "-1")
    End Sub
End Module
