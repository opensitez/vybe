' vybe-test: vb/vb_system_collection_api_matrix/collection_api_matrix_set_intersection_like
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
        Dim left As New HashSet(Of Integer)({1, 2, 3, 4})
        Dim right As New HashSet(Of Integer)({3, 4, 5})
        left.IntersectWith(right)

        Dim ordered As New List(Of Integer)(left)
        ordered.Sort()

        __Check(CStr(String.Join(",", ordered)), "3,4")
        __Check(CStr(left.Count), "2")
    End Sub
End Module
