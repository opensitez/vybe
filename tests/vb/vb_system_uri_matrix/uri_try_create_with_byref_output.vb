' vybe-test: vb/vb_system_uri_matrix/uri_try_create_with_byref_output
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
        Dim parsed As Uri = Nothing
        Dim ok As Boolean = Uri.TryCreate("https://example.com", UriKind.Absolute, parsed)
        __Check(CStr(ok), "True")
        __Check(CStr(parsed IsNot Nothing), "True")
    End Sub
End Module
