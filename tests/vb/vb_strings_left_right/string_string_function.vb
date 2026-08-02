' vybe-test: vb/vb_strings_left_right/string_string_function
' origin: languages/vb/tests/vb/test_vb_strings_left_right.rs

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
        ' Creates a string of a repeated character
        __Check(CStr(String(5, "x"c)), "xxxxx")
        ' Also works with char code
        __Check(CStr(String(3, 42)), "***") ' 42 is '*'
    End Sub
End Module
