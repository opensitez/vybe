' vybe-test: vb/vb_char_unicode_category_ascii_checks/test_vb_char_string_indexed_character_checks
' origin: languages/vb/tests/vb/test_vb_char_unicode_category_ascii_checks.rs

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

Imports System

Module Program
    Sub Main()
        Dim text = "A1 #"
        __Check(CStr(Char.IsLetter(text, 0) & "|" & Char.IsDigit(text, 1) & "|" & Char.IsWhiteSpace(text, 2) & "|" & Char.IsPunctuation(text, 3)), "True|True|True|True")
    End Sub
End Module
