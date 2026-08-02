' vybe-test: vb/vb_dictionary_key_collection/test_vb_dict_remove_key_value_pair
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
        Dim dict As New Dictionary(Of Integer, String) From {{1, "One"}, {2, "Two"}}
        Dim removedWrong As Boolean = dict.Remove(New KeyValuePair(Of Integer, String)(1, "Wrong"))
        Dim removedRight As Boolean = dict.Remove(New KeyValuePair(Of Integer, String)(1, "One"))
        __Check(CStr(removedWrong), "False")
        __Check(CStr(removedRight), "True")
        __Check(CStr(dict.Count), "1")
    End Sub
End Module
