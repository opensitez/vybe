' vybe-test: vb/vb_tuple_hashset_dictionary_keys/test_vb_tuple_dictionary_try_get_value
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
        Dim dict As New Dictionary(Of (Integer, Integer), String)()
        dict((10, 20)) = "PointA"
        Dim res As String = Nothing
        Dim found = dict.TryGetValue((10, 20), res)
        __Check(CStr(found & ":" & res), "True:PointA")
    End Sub
End Module
