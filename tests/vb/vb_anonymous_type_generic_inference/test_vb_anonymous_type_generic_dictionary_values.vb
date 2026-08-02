' vybe-test: vb/vb_anonymous_type_generic_inference/test_vb_anonymous_type_generic_dictionary_values
' origin: languages/vb/tests/vb/test_vb_anonymous_type_generic_inference.rs

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
    Private Function CreateDict(Of TKey, TValue)(k As TKey, v As TValue) As Dictionary(Of TKey, TValue)
        Dim d As New Dictionary(Of TKey, TValue)()
        d(k) = v
        Return d
    End Function

    Sub Main()
        Dim valObj = New With {.Status = "OK"}
        Dim dict = CreateDict("Key1", valObj)
        __Check(CStr(dict("Key1").Status), "OK")
    End Sub
End Module
