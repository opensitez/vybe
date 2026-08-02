' vybe-test: vb/vb_string_composite_formatting_alignment/test_vb_string_format_date_time_format_specifiers
' origin: languages/vb/tests/vb/test_vb_string_composite_formatting_alignment.rs

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
        Dim dt As New DateTime(2025, 4, 1, 14, 5, 9)
        Dim res = String.Format(CultureInfo.InvariantCulture, "{0:yyyy-MM-dd HH:mm:ss}", dt)
        __Check(CStr(res), "2025-04-01 14:05:09")
    End Sub
End Module
