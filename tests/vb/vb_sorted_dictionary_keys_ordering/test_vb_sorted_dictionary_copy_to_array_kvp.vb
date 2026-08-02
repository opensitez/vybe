' vybe-test: vb/vb_sorted_dictionary_keys_ordering/test_vb_sorted_dictionary_copy_to_array_kvp
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

Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of String, Integer)()
        dict("X") = 10
        dict("Y") = 20

        Dim array(1) As KeyValuePair(Of String, Integer)
        dict.CopyTo(array, 0)
        __Check(CStr(array(0).Key & "=" & array(0).Value & "|" & array(1).Key & "=" & array(1).Value), "X=10|Y=20")
    End Sub
End Module
