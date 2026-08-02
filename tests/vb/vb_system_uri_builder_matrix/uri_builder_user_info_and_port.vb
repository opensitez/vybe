' vybe-test: vb/vb_system_uri_builder_matrix/uri_builder_user_info_and_port
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
        builder.UserName = "alice"
        builder.Password = "secret"
        builder.Port = 8443

        Dim built As Uri = builder.Uri
        __Check(CStr(built.UserInfo), "alice:secret")
        __Check(CStr(built.Authority), "alice:secret@example.com:8443")
        __Check(CStr(built.Port), "8443")
    End Sub
End Module
