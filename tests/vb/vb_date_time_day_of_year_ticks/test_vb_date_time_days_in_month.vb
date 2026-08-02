' vybe-test: vb/vb_date_time_day_of_year_ticks/test_vb_date_time_days_in_month
' origin: languages/vb/tests/vb/test_vb_date_time_day_of_year_ticks.rs

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
        Dim feb2024 = DateTime.DaysInMonth(2024, 2)
        Dim feb2025 = DateTime.DaysInMonth(2025, 2)
        __Check(CStr(feb2024 & "|" & feb2025), "29|28")
    End Sub
End Module
