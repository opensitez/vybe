' vybe-test: vb/vb_list_remove_all_predicate/test_vb_list_clear_resets_count_preserves_capacity
' origin: languages/vb/tests/vb/test_vb_list_remove_all_predicate.rs

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

Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim list As New List(Of Integer)(50) From {1, 2, 3, 4, 5}
        list.Clear()
        __Check(CStr(list.Count & "|" & (list.Capacity >= 50)), "0|True")
    End Sub
End Module
