' vybe-test: vb/vb_encoding_utf8_ascii_bytes/test_vb_encoding_get_preamble_utf8_bom
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
        Dim utf8Bom = Encoding.UTF8.GetPreamble()
        ' UTF-8 BOM is 239, 187, 191 (EF BB BF)
        __Check(CStr(String.Join(",", utf8Bom)), "239,187,191")
    End Sub
End Module
