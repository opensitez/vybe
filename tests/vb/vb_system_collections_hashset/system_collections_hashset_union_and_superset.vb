' vybe-test: vb/vb_system_collections_hashset/system_collections_hashset_union_and_superset
' origin: languages/vb/tests/vb/test_vb_system_collections_hashset.rs

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
        Dim left As New HashSet(Of Integer)()
        left.Add(1)
        left.Add(3)

        Dim right As New HashSet(Of Integer)()
        right.Add(3)
        right.Add(4)

        left.UnionWith(right)
        __Check(CStr(left.Count), "3")
        __Check(CStr(left.Contains(4)), "True")
        __Check(CStr(left.IsSupersetOf(right)), "True")
    End Sub
End Module
