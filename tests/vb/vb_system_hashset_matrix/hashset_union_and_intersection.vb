' vybe-test: vb/vb_system_hashset_matrix/hashset_union_and_intersection
' origin: languages/vb/tests/vb/test_vb_system_hashset_matrix.rs

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
        Dim left As New HashSet(Of Integer) From {1, 2}
        Dim right As New HashSet(Of Integer) From {2, 3}
        left.UnionWith(right)
        __Check(CStr(left.Count), "3")
        Dim intersection As New HashSet(Of Integer)({1, 2, 3})
        intersection.IntersectWith(New HashSet(Of Integer)({2, 3, 4}))
        __Check(CStr(intersection.Count), "2")
    End Sub
End Module
