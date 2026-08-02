' vybe-test: vb/vb_system_uri_builder_matrix/uri_builder_empty_host_becomes_relative_uri
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
        builder.Path = "relative/path"
        builder.Query = "x=1"

        __Check(CStr(builder.Uri.IsAbsoluteUri), "False")
        __Check(CStr(builder.Uri.PathAndQuery), "relative/path?x=1")
        __Check(CStr(builder.Uri.ToString()), "relative/path?x=1")
    End Sub
End Module
