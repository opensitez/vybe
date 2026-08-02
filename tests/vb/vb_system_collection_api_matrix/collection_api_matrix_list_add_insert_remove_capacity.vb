' vybe-test: vb/vb_system_collection_api_matrix/collection_api_matrix_list_add_insert_remove_capacity
' origin: languages/vb/tests/vb/test_vb_system_collection_api_matrix.rs

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
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim list As New List(Of Integer) From {1, 2, 4}
        list.Insert(2, 3)
        list.RemoveAt(0)

        __Check(CStr(list.Count), "3")
        __Check(CStr(list(0)), "2")
        __Check(CStr(list(2)), "4")
    End Sub
End Module
