' vybe-test: vb/vb_sorted_dictionary_keys_ordering/test_vb_sorted_dictionary_linq_query_projections
' origin: languages/vb/tests/vb/test_vb_sorted_dictionary_keys_ordering.rs

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
Imports System.Linq

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of Integer, Integer)()
        dict(1) = 10
        dict(2) = 20
        dict(3) = 30

        Dim query = dict.Where(Function(kvp) kvp.Value > 15).Select(Function(kvp) kvp.Key)
        __Check(CStr(String.Join(",", query)), "2,3")
    End Sub
End Module
