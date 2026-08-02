' vybe-test: vb/vb_double_try_parse_cultures/test_vb_double_to_string_culture_formatting_roundtrip
' origin: languages/vb/tests/vb/test_vb_double_try_parse_cultures.rs

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
Imports System.Globalization

Module Program
    Sub Main()
        Dim orig As Double = 123456.789
        Dim formatted = orig.ToString("N3", CultureInfo.InvariantCulture)
        Dim restored As Double
        Dim ok = Double.TryParse(formatted, NumberStyles.Number, CultureInfo.InvariantCulture, restored)
        __Check(CStr(formatted & "|" & (orig = restored)), "123,456.789|True")
    End Sub
End Module
