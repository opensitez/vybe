' vybe-test: vb/vb_date_time_parse_exact_formats/test_vb_date_time_try_parse_custom_culture_german
' origin: languages/vb/tests/vb/test_vb_date_time_parse_exact_formats.rs

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
        Dim culture = CultureInfo.GetCultureInfo("de-DE")
        Dim dt = DateTime.Parse("15.08.2025", culture)
        __Check(CStr(dt.Day & "." & dt.Month & "." & dt.Year), "15.8.2025")
    End Sub
End Module
