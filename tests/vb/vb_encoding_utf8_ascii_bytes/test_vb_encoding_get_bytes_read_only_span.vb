' vybe-test: vb/vb_encoding_utf8_ascii_bytes/test_vb_encoding_get_bytes_read_only_span
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

Imports System
Imports System.Text

Module Program
    Sub Main()
        Dim textSpan As ReadOnlySpan(Of Char) = "SpanEncoding".ToCharArray()
        Dim bytes = Encoding.UTF8.GetBytes(textSpan)
        __Check(CStr(Encoding.UTF8.GetString(bytes)), "SpanEncoding")
    End Sub
End Module
