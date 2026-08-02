' vybe-test: vb/vb_sorted_dictionary_operations/test_vb_sorted_dict_remove_key
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
        Dim dict As New SortedDictionary(Of Integer, String) From {{1, "A"}, {2, "B"}}
        Dim ok As Boolean = dict.Remove(1)
        __Check(CStr(ok), "True")
        __Check(CStr(dict.Count), "1")
    End Sub
End Module
