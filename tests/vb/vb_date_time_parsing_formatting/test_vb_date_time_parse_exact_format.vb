' vybe-test: vb/vb_date_time_parsing_formatting/test_vb_date_time_parse_exact_format
' origin: languages/vb/tests/vb/test_vb_date_time_parsing_formatting.rs

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
        Dim strVal As String = "2025-05-15 14:30:00"
        Dim dt As DateTime = DateTime.ParseExact(strVal, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)
        __Check(CStr(dt.Year & "-" & dt.Month & "-" & dt.Day & " " & dt.Hour & ":" & dt.Minute), "2025-5-15 14:30")
    End Sub
End Module
