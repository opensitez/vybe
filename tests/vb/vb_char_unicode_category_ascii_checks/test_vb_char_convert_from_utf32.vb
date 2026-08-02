' vybe-test: vb/vb_char_unicode_category_ascii_checks/test_vb_char_convert_from_utf32
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
        Dim str = Char.ConvertFromUtf32(&H1F600)
        __Check(CStr(str.Length & "|" & Char.IsSurrogatePair(str, 0)), "2|True")
    End Sub
End Module
