' vybe-test: vb/vb_hashset_set_algebra/hashset_set_equals_and_subset_checks
' origin: languages/vb/tests/vb/test_vb_hashset_set_algebra.rs

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

        Dim c As New HashSet(Of Integer)()
        c.Add(1)

        __Check(CStr(a.SetEquals(b)), "True")
        __Check(CStr(c.IsSubsetOf(a)), "True")
        __Check(CStr(a.IsSupersetOf(c)), "True")
    End Sub
End Module
