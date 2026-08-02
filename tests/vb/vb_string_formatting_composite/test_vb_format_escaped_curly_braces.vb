' vybe-test: vb/vb_string_formatting_composite/test_vb_format_escaped_curly_braces
' origin: languages/vb/tests/vb/test_vb_string_formatting_composite.rs

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
        Dim val As Integer = 42
        __Check(CStr(String.Format("{{Value: {0}}}", val)), "{Value: 42}")
    End Sub
End Module
