' vybe-test: vb/vb_encoding_utf8_ascii_bytes/test_vb_encoding_get_char_count_buffer
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
        Dim bytes As Byte() = Encoding.UTF8.GetBytes("VisualBasic")
        Dim charCount = Encoding.UTF8.GetCharCount(bytes, 0, 6)
        __Check(CStr(charCount), "6")
    End Sub
End Module
