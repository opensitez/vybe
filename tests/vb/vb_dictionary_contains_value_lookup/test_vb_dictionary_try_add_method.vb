' vybe-test: vb/vb_dictionary_contains_value_lookup/test_vb_dictionary_try_add_method
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
        Dim dict As New Dictionary(Of String, Integer)()
        Dim added1 As Boolean = dict.TryAdd("Key", 100)
        Dim added2 As Boolean = dict.TryAdd("Key", 200)
        __Check(CStr(added1 & "|" & added2 & "|" & dict("Key")), "True|False|100")
    End Sub
End Module
