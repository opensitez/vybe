' vybe-test: vb/vb_dictionary_init_nested/dictionary_init_nested
' origin: languages/vb/tests/vb/test_vb_dictionary_init_nested.rs

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
        Dim dict As New Dictionary(Of String, Object) From {
            {"A", New With {.Value = 1}},
            {"B", New With {.Value = 2}}
        }
        
        ' Late binding used to access .Value
        __Check(CStr(dict("A").Value), "1")
        __Check(CStr(dict("B").Value), "2")
    End Sub
End Module
