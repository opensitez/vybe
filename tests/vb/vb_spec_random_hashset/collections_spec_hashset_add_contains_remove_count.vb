' vybe-test: vb/vb_spec_random_hashset/collections_spec_hashset_add_contains_remove_count
' origin: languages/vb/tests/vb/test_vb_spec_random_hashset.rs

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
Module Program
    Sub Main()
        Dim items As New HashSet(Of Integer)()
        items.Add(1)
        items.Add(2)
        items.Add(2)
        __Check(CStr(items.Count), "2")
        __Check(CStr(items.Contains(2)), "True")
        items.Remove(1)
        __Check(CStr(items.Contains(1)), "False")
    End Sub
End Module
