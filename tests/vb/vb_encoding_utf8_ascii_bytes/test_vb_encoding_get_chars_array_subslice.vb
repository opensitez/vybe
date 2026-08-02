' vybe-test: vb/vb_encoding_utf8_ascii_bytes/test_vb_encoding_get_chars_array_subslice
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
        Dim chars(5) As Char
        Dim count = Encoding.UTF8.GetChars(bytes, 0, 6, chars, 0)
        __Check(CStr(count & ":" & New String(chars)), "6:Visual")
    End Sub
End Module
