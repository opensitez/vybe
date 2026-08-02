' vybe-test: vb/vb_system_collections_hashset/system_collections_hashset_intersect_and_except
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
        Dim baseSet As New HashSet(Of Integer)()
        baseSet.Add(1)
        baseSet.Add(2)
        baseSet.Add(3)

        Dim intersectWith As New HashSet(Of Integer)()
        intersectWith.Add(2)
        intersectWith.Add(3)
        intersectWith.Add(4)

        baseSet.IntersectWith(intersectWith)
        __Check(CStr(baseSet.Count), "2")
        __Check(CStr(baseSet.Contains(2)), "True")

        baseSet.ExceptWith(intersectWith)
        __Check(CStr(baseSet.Count), "0")
    End Sub
End Module
