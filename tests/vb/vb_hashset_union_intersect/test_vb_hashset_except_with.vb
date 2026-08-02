' vybe-test: vb/vb_hashset_union_intersect/test_vb_hashset_except_with
' origin: languages/vb/tests/vb/test_vb_hashset_union_intersect.rs

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
        Dim set1 As New HashSet(Of Integer) From {1, 2, 3, 4}
        Dim set2 As New HashSet(Of Integer) From {2, 4}
        set1.ExceptWith(set2)
        __Check(CStr(String.Join(",", set1)), "1,3")
    End Sub
End Module
