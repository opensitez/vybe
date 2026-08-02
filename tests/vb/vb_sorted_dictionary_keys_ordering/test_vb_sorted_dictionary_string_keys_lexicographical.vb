' vybe-test: vb/vb_sorted_dictionary_keys_ordering/test_vb_sorted_dictionary_string_keys_lexicographical
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

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of String, Integer)()
        dict("Banana") = 2
        dict("Apple") = 1
        dict("Cherry") = 3

        Dim keys = String.Join(",", dict.Keys)
        __Check(CStr(keys), "Apple,Banana,Cherry")
    End Sub
End Module
