' vybe-test: vb/vb_date_time_parsing_formatting/test_vb_date_time_try_parse_success_and_failure
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

Module Program
    Sub Main()
        Dim dt As DateTime
        Dim ok1 As Boolean = DateTime.TryParse("2025-01-01", dt)
        Dim ok2 As Boolean = DateTime.TryParse("invalid date", dt)
        __Check(CStr(ok1), "True")
        __Check(CStr(ok2), "False")
    End Sub
End Module
