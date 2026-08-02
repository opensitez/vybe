' vybe-test: vb/vb_system_encoding_matrix/encoding_unicode_uses_two_bytes_for_ascii_char
' origin: languages/vb/tests/vb/test_vb_system_encoding_matrix.rs

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

Module M
    Sub Main()
        Dim bytes() As Byte = Encoding.Unicode.GetBytes("x")
        __Check(CStr(bytes.Length), "2")
        __Check(CStr(Encoding.Unicode.GetString(bytes)), "x")
    End Sub
End Module
