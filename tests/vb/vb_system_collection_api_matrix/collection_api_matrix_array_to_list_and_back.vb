' vybe-test: vb/vb_system_collection_api_matrix/collection_api_matrix_array_to_list_and_back
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
        Dim source As Integer() = {1, 2, 3}
        Dim list As New List(Of Integer)(source)
        list.Add(4)
        Dim arr As Integer() = list.ToArray()

        __Check(CStr(list.Count), "4")
        __Check(CStr(arr.Length), "4")
        __Check(CStr(arr(3)), "4")
    End Sub
End Module
