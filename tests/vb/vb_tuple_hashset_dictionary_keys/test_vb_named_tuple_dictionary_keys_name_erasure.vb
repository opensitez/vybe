' vybe-test: vb/vb_tuple_hashset_dictionary_keys/test_vb_named_tuple_dictionary_keys_name_erasure
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
        Dim dict As New Dictionary(Of (X As Integer, Y As Integer), String)()
        dict((10, 20)) = "Location1"

        Dim searchTuple As (Col As Integer, Row As Integer) = (10, 20)
        __Check(CStr(dict(searchTuple)), "Location1")
    End Sub
End Module
