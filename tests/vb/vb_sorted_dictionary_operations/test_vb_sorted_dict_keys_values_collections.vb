' vybe-test: vb/vb_sorted_dictionary_operations/test_vb_sorted_dict_keys_values_collections
' origin: languages/vb/tests/vb/test_vb_sorted_dictionary_operations.rs

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
        Dim dict As New SortedDictionary(Of Integer, String) From {{3, "C"}, {1, "A"}, {2, "B"}}
        __Check(CStr(String.Join(",", dict.Keys)), "1,2,3")
        __Check(CStr(String.Join(",", dict.Values)), "A,B,C")
    End Sub
End Module
