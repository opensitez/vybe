' vybe-test: vb/vb_sorted_dictionary_keys_ordering/test_vb_sorted_dictionary_duplicate_key_add_throws_argument_exception
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

Imports System
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of String, Integer)()
        dict.Add("Unique", 1)
        Try
            dict.Add("Unique", 2)
        Catch ex As ArgumentException
            __Check(CStr("ArgumentException Caught on Duplicate Key Add"), "ArgumentException Caught on Duplicate Key Add")
        End Try
    End Sub
End Module
