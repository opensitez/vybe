' vybe-test: vb/vb_module_alias_imports/module_alias_text_encoding_ascii
' origin: languages/vb/tests/vb/test_vb_module_alias_imports.rs

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

Imports EncodingAlias = System.Text.Encoding

Module M
    Sub Main()
        Dim bytes() As Byte = EncodingAlias.UTF8.GetBytes("go")
        __Check(CStr(bytes.Length), "2")
        __Check(CStr(CStr(bytes(0))), "103")
    End Sub
End Module
