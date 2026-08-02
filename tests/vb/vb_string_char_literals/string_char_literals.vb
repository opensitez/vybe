' vybe-test: vb/vb_string_char_literals/string_char_literals
' origin: languages/vb/tests/vb/test_vb_string_char_literals.rs

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

Module M
    Sub Main()
        ' Double quotes is a String
        Dim s As String = "A"
        
        ' Double quotes followed by c is a Char
        Dim c As Char = "A"c
        
        __Check(CStr(s.GetType().Name), "String")
        __Check(CStr(c.GetType().Name), "Char")
        
        ' Escaping double quotes inside a string literal
        Dim q As String = "He said ""Hello"" to me"
        __Check(CStr(q), "He said ""Hello"" to me")
    End Sub
End Module
