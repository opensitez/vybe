' vybe-test: vb/vb_sorted_dictionary_keys_ordering/test_vb_sorted_dictionary_null_key_throws_argument_null
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
        Try
            dict.Add(Nothing, 100)
        Catch ex As ArgumentNullException
            __Check(CStr("ArgumentNullException Caught on Null Key"), "ArgumentNullException Caught on Null Key")
        End Try
    End Sub
End Module
