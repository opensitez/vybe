' vybe-test: vb/vb_system_uri_builder_matrix/uri_builder_populates_uri_properties
' origin: languages/vb/tests/vb/test_vb_system_uri_builder_matrix.rs

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

Module M
    Sub Main()
        Dim builder As New UriBuilder()
        builder.Scheme = "https"
        builder.Host = "example.com"
        builder.Port = 443
        builder.Path = "/api/v1"
        builder.Query = "page=1"
        builder.Fragment = "section"

        __Check(CStr(builder.Uri.Scheme), "https")
        __Check(CStr(builder.Uri.Host), "example.com")
        __Check(CStr(builder.Uri.Port), "443")
        __Check(CStr(builder.Uri.AbsolutePath), "/api/v1")
        __Check(CStr(builder.Uri.Query), "?page=1")
        __Check(CStr(builder.Uri.Fragment), "#section")
    End Sub
End Module
