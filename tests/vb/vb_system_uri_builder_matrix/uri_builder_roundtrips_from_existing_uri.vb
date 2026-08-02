' vybe-test: vb/vb_system_uri_builder_matrix/uri_builder_roundtrips_from_existing_uri
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
        Dim original As New Uri("https://example.com/blog/index.html?x=1#top")
        Dim builder As New UriBuilder(original)

        __Check(CStr(builder.Uri.AbsoluteUri), "https://example.com/blog/index.html?x=1#top")
        __Check(CStr(builder.Path), "/blog/index.html")
        __Check(CStr(builder.Query), "?x=1")
        __Check(CStr(builder.Fragment), "#top")
    End Sub
End Module
