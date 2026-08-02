' vybe-test: vb/vb_sorted_dictionary_keys_ordering/test_vb_sorted_dictionary_try_get_value
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
        dict("Score") = 100

        Dim val As Integer
        Dim found = dict.TryGetValue("Score", val)
        Dim missing = dict.TryGetValue("Missing", val)

        __Check(CStr(found & "|" & val & "|" & missing), "True|100|False")
    End Sub
End Module
