' vybe-test: vb/vb_system_collections_hashset/system_collections_hashset_set_relations
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
        Dim a As New HashSet(Of Integer)()
        a.Add(1)
        a.Add(2)

        Dim b As New HashSet(Of Integer)()
        b.Add(1)
        b.Add(2)

        __Check(CStr(a.SetEquals(b)), "True")
        __Check(CStr(a.Overlaps(b)), "True")
        __Check(CStr(a.IsProperSubsetOf(b)), "False")
    End Sub
End Module
