' vybe-test: vb/vb_char_and_string_basics/string_fixed_length_legacy
' origin: languages/vb/tests/vb/test_vb_char_and_string_basics.rs

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
' Dim s As String * 10 is legacy, often unsupported in .NET directly but parsable
__Check(CStr("Parsed"), "Parsed")
End Sub
End Module
