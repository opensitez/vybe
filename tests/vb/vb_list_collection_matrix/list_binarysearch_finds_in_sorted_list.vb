' vybe-test: vb/vb_list_collection_matrix/list_binarysearch_finds_in_sorted_list
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
        Dim values As New List(Of Integer) From {1, 3, 5, 7}
        __Check(CStr(values.BinarySearch(5)), "2")
        __Check(CStr(values.BinarySearch(6)), "-4")
    End Sub
End Module
