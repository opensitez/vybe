' vybe-test: vb/vb_uri_builder_query_parsing/test_vb_uri_builder_construct_url
' origin: languages/vb/tests/vb/test_vb_uri_builder_query_parsing.rs

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
        Dim builder As New UriBuilder("https", "example.com", 8080, "api/v1")
        builder.Query = "key=value"
        Dim uri As Uri = builder.Uri
        __Check(CStr(uri.Scheme), "https")
        __Check(CStr(uri.Host), "example.com")
        __Check(CStr(uri.Port), "8080")
        __Check(CStr(uri.Query), "?key=value")
    End Sub
End Module
