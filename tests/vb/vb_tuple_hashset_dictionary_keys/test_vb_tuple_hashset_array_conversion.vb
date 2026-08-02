' vybe-test: vb/vb_tuple_hashset_dictionary_keys/test_vb_tuple_hashset_array_conversion
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
Imports System.Linq

Module Program
    Sub Main()
        Dim set As New HashSet(Of (Integer, String)) From {(1, "A"), (2, "B")}
        Dim arr = set.ToArray()
        __Check(CStr(arr.Length & ":" & arr(0).Item2 & "," & arr(1).Item2), "2:A,B")
    End Sub
End Module
