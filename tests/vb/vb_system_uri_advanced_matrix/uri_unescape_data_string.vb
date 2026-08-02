' vybe-test: vb/vb_system_uri_advanced_matrix/uri_unescape_data_string
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
        Dim encoded As String = Uri.EscapeDataString("a b")
        Dim decoded As String = Uri.UnescapeDataString(encoded)

        __Check(CStr(encoded), "a%20b")
        __Check(CStr(decoded), "a b")
    End Sub
End Module
