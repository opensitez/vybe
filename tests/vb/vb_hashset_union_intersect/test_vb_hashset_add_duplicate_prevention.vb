' vybe-test: vb/vb_hashset_union_intersect/test_vb_hashset_add_duplicate_prevention
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
        Dim hs As New HashSet(Of Integer)()
        Dim added1 As Boolean = hs.Add(10)
        Dim added2 As Boolean = hs.Add(10)
        __Check(CStr(added1), "True")
        __Check(CStr(added2), "False")
        __Check(CStr(hs.Count), "1")
    End Sub
End Module
