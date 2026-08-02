' vybe-test: vb/vb_sorted_dictionary_custom_key_comparer/test_vb_sorted_dictionary_default_key_ordering
' origin: languages/vb/tests/vb/test_vb_sorted_dictionary_custom_key_comparer.rs

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
        Dim dict As New SortedDictionary(Of Integer, String)()
        dict(30) = "Thirty"
        dict(10) = "Ten"
        dict(20) = "Twenty"
        __Check(CStr(String.Join(",", dict.Keys)), "10,20,30")
    End Sub
End Module
