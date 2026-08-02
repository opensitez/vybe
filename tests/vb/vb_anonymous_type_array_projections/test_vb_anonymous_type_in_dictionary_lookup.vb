' vybe-test: vb/vb_anonymous_type_array_projections/test_vb_anonymous_type_in_dictionary_lookup
' origin: languages/vb/tests/vb/test_vb_anonymous_type_array_projections.rs

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
        Dim dict As New Dictionary(Of Object, String)()
        Dim keyObj = New With {Key .ID = 42}
        dict(keyObj) = "FoundData"
        Dim lookupObj = New With {Key .ID = 42}
        __Check(CStr(dict(lookupObj)), "FoundData")
    End Sub
End Module
