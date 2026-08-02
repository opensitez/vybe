' vybe-test: vb/vb_system_uri_advanced_matrix/uri_from_string_to_constructor_roundtrip
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
        Dim original As Uri = New Uri("https://example.com/blog/index.html")
        Dim parsed As Uri = New Uri(original.ToString())
        __Check(CStr(parsed = original), "True")
        __Check(CStr(parsed.AbsoluteUri), "https://example.com/blog/index.html")
    End Sub
End Module
