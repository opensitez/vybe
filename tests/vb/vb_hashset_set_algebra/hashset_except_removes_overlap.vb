' vybe-test: vb/vb_hashset_set_algebra/hashset_except_removes_overlap
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
        Dim baseSet As New HashSet(Of Integer)()
        baseSet.Add(1)
        baseSet.Add(2)
        baseSet.Add(3)

        Dim toRemove As New HashSet(Of Integer)()
        toRemove.Add(2)

        baseSet.ExceptWith(toRemove)
        __Check(CStr(baseSet.Count), "2")
        __Check(CStr(baseSet.Contains(2)), "False")
    End Sub
End Module
