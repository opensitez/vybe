' vybe-test: vb/vb_system_uri_advanced_matrix/uri_is_absolute_or_relative
' origin: languages/vb/tests/vb/test_vb_system_uri_advanced_matrix.rs

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
        Dim absoluteUri As New Uri("https://example.com/api")
        Dim relativeUri As New Uri("/api/status", UriKind.Relative)

        __Check(CStr(absoluteUri.IsAbsoluteUri), "True")
        __Check(CStr(relativeUri.IsAbsoluteUri), "False")
    End Sub
End Module
