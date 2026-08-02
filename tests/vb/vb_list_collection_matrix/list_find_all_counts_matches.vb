' vybe-test: vb/vb_list_collection_matrix/list_find_all_counts_matches
' origin: languages/vb/tests/vb/test_vb_list_collection_matrix.rs

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

Module M
    Sub Main()
        Dim values As New List(Of Integer) From {1, 2, 3, 4}
        Dim evens As List(Of Integer) = values.FindAll(Function(value As Integer) value Mod 2 = 0)
        __Check(CStr(evens.Count), "2")
        __Check(CStr(evens(0)), "2")
        __Check(CStr(evens(1)), "4")
    End Sub
End Module
