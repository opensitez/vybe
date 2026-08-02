' vybe-test: vb/vb_dictionary_key_collection/test_vb_dict_try_add_method
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
        Dim dict As New Dictionary(Of Integer, String)
        Dim firstAdd As Boolean = dict.TryAdd(1, "First")
        Dim secondAdd As Boolean = dict.TryAdd(1, "Second")
        __Check(CStr(firstAdd), "True")
        __Check(CStr(secondAdd), "False")
        __Check(CStr(dict(1)), "First")
    End Sub
End Module
