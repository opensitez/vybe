' vybe-test: vb/vb_tuple_hashset_dictionary_keys/test_vb_tuple_hashset_intersect_with
' origin: languages/vb/tests/vb/test_vb_tuple_hashset_dictionary_keys.rs

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
        Dim set1 As New HashSet(Of (Integer, String)) From {(1, "A"), (2, "B")}
        Dim set2 As New HashSet(Of (Integer, String)) From {(2, "B"), (3, "C")}
        set1.IntersectWith(set2)
        __Check(CStr(set1.Count & ":" & set1.First().Item2), "1:B")
    End Sub
End Module
