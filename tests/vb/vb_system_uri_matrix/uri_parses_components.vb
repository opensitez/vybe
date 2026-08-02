' vybe-test: vb/vb_system_uri_matrix/uri_parses_components
' origin: languages/vb/tests/vb/test_vb_system_uri_matrix.rs

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
        Dim u As New Uri("https://example.com:8443/app/index.html?x=1#top")
        __Check(CStr(u.Scheme), "https")
        __Check(CStr(u.Host), "example.com")
        __Check(CStr(u.Port), "8443")
        __Check(CStr(u.AbsolutePath), "/app/index.html")
        __Check(CStr(u.Query), "?x=1")
        __Check(CStr(u.Fragment), "#top")
    End Sub
End Module
