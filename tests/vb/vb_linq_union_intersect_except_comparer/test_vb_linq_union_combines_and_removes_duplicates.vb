' vybe-test: vb/vb_linq_union_intersect_except_comparer/test_vb_linq_union_combines_and_removes_duplicates
' origin: languages/vb/tests/vb/test_vb_linq_union_intersect_except_comparer.rs

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

Imports System.Linq

Module Program
    Sub Main()
        Dim seq1 = {1, 2, 3}
        Dim seq2 = {3, 4, 5}
        Dim res = seq1.Union(seq2)
        __Check(CStr(String.Join(",", res)), "1,2,3,4,5")
    End Sub
End Module
