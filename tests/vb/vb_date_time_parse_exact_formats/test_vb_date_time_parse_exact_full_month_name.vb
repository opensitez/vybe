' vybe-test: vb/vb_date_time_parse_exact_formats/test_vb_date_time_parse_exact_full_month_name
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
        Dim dt = DateTime.ParseExact("August 15, 2025", "MMMM dd, yyyy", CultureInfo.InvariantCulture)
        __Check(CStr(dt.Month & "/" & dt.Day), "8/15")
    End Sub
End Module
