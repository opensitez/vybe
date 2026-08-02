' vybe-test: vb/vb_dictionary_contains_value_lookup/test_vb_dictionary_keys_collection_iteration
' origin: languages/vb/tests/vb/test_vb_dictionary_contains_value_lookup.rs

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
        Dim dict As New Dictionary(Of String, Integer) From {{"X", 1}, {"Y", 2}}
        Dim keys As String = String.Join(",", dict.Keys)
        __Check(CStr(keys), "X,Y")
    End Sub
End Module
