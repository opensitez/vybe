' vybe-test: vb/vb_dictionary_contains_value_lookup/test_vb_dictionary_remove_key_with_out_value_parameter
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
        Dim dict As New Dictionary(Of String, Integer) From {{"Key", 42}}
        Dim removedVal As Integer
        Dim removed As Boolean = dict.Remove("Key", removedVal)
        __Check(CStr(removed & ":" & removedVal), "True:42")
    End Sub
End Module
