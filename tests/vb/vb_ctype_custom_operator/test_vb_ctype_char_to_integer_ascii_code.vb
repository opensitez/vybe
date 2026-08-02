' vybe-test: vb/vb_ctype_custom_operator/test_vb_ctype_char_to_integer_ascii_code
' origin: languages/vb/tests/vb/test_vb_ctype_custom_operator.rs

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

Module Program
    Sub Main()
        Dim ch As Char = "A"c
        Dim code As Integer = CType(ch, Integer)
        Dim restored As Char = CType(code, Char)
        __Check(CStr(code & "|" & restored), "65|A")
    End Sub
End Module
