' vybe-test: vb/vb_convert_change_type_reflection/test_vb_convert_change_type_uri_from_string
' origin: languages/vb/tests/vb/test_vb_convert_change_type_reflection.rs

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

Module Program
    Sub Main()
        Dim urlStr = "https://example.com/api"
        Dim u As New Uri(urlStr)
        __Check(CStr(u.Host & "|" & u.Scheme), "example.com|https")
    End Sub
End Module
