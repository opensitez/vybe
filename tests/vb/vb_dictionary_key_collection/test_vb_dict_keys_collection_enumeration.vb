' vybe-test: vb/vb_dictionary_key_collection/test_vb_dict_keys_collection_enumeration
' origin: languages/vb/tests/vb/test_vb_dictionary_key_collection.rs

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
        Dim dict As New Dictionary(Of String, Integer) From {{"A", 1}, {"B", 2}, {"C", 3}}
        Dim keys As Dictionary(Of String, Integer).KeyCollection = dict.Keys
        __Check(CStr(keys.Count), "3")
        __Check(CStr(String.Join(",", keys)), "A,B,C")
    End Sub
End Module
