' vybe-test: vb/vb_collection_initializers_dict/collection_initializer_dictionary
' origin: languages/vb/tests/vb/test_vb_collection_initializers_dict.rs

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

Module M
    Sub Main()
        ' Dictionary collection initializer syntax (uses nested braces)
        Dim dict As New Dictionary(Of String, Integer) From {
            {"A", 1},
            {"B", 2},
            {"C", 3}
        }
        
        __Check(CStr(dict.Count), "3")
        __Check(CStr(dict("B")), "2")
    End Sub
End Module
