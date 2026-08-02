' vybe-test: vb/vb_char_unicode_category_ascii_checks/test_vb_char_convert_to_utf32_surrogate_pair
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
        Dim highSurrogate = ChrW(&HD83D)
        Dim lowSurrogate = ChrW(&HDE00)
        Dim utf32 = Char.ConvertToUtf32(highSurrogate, lowSurrogate)
        __Check(CStr(Hex(utf32)), "1F600")
    End Sub
End Module
