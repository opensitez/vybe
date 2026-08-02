' vybe-test: vb/vb_encoding_utf8_ascii_bytes/test_vb_encoding_clone_customization
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
        Dim enc As Encoding = CType(Encoding.UTF8.Clone(), Encoding)
        __Check(CStr(enc.IsReadOnly), "False")
    End Sub
End Module
