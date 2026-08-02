' vybe-test: vb/vb_sorted_dictionary_keys_ordering/test_vb_sorted_dictionary_struct_value_type
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

Structure ColorRGB
    Public R, G, B As Byte
End Structure

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of String, ColorRGB)()
        dict("Red") = New ColorRGB With {.R = 255, .G = 0, .B = 0}
        dict("Green") = New ColorRGB With {.R = 0, .G = 255, .B = 0}

        __Check(CStr(dict("Green").G), "255")
    End Sub
End Module
