' vybe-test: vb/vb_string_trim_start_end_chars/test_vb_string_trim_non_matching_char
' origin: languages/vb/tests/vb/test_vb_string_trim_start_end_chars.rs

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
        Dim s As String = "Hello"
        __Check(CStr(s.Trim("x"c)), "Hello")
    End Sub
End Module
