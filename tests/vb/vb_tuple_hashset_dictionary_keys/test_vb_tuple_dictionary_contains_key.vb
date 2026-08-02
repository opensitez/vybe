' vybe-test: vb/vb_tuple_hashset_dictionary_keys/test_vb_tuple_dictionary_contains_key
' origin: languages/vb/tests/vb/test_vb_tuple_hashset_dictionary_keys.rs

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
        Dim dict As New Dictionary(Of (String, Integer), Boolean)()
        dict(("Admin", 1)) = True
        __Check(CStr(dict.ContainsKey(("Admin", 1)) & "|" & dict.ContainsKey(("Guest", 2))), "True|False")
    End Sub
End Module
