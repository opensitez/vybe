' vybe-test: vb/vb_sorted_dictionary_keys_ordering/test_vb_sorted_dictionary_constructor_existing_dictionary
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
        Dim rawDict As New Dictionary(Of Integer, String)()
        rawDict(3) = "C"
        rawDict(1) = "A"
        rawDict(2) = "B"

        Dim sorted As New SortedDictionary(Of Integer, String)(rawDict)
        __Check(CStr(String.Join(",", sorted.Keys)), "1,2,3")
    End Sub
End Module
