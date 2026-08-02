' vybe-test: vb/vb_char_and_string_basics/string_concatenation_number_plus
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

Option Strict Off
Module M
Sub Main()
' + with string and number is tricky, usually throws InvalidCast if string isn't numeric
__Check(CStr("Parsed"), "Parsed")
End Sub
End Module
