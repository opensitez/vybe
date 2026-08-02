' vybe-test: vb/vb_encoding_utf8_ascii_bytes/test_vb_encoding_ascii_get_bytes_replaces_non_ascii
' origin: languages/vb/tests/vb/test_vb_encoding_utf8_ascii_bytes.rs

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

Imports System.Text

Module Program
    Sub Main()
        ' ASCII replaces non-ASCII characters with '?' (63)
        Dim text = "Hello World"
        Dim bytes = Encoding.ASCII.GetBytes(text)
        Dim restored = Encoding.ASCII.GetString(bytes)
        __Check(CStr(restored), "Hello World")
    End Sub
End Module
