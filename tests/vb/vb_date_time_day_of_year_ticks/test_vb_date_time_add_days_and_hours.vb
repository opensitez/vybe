' vybe-test: vb/vb_date_time_day_of_year_ticks/test_vb_date_time_add_days_and_hours
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
        Dim dt As New DateTime(2025, 1, 1, 10, 0, 0)
        Dim dtFuture = dt.AddDays(5).AddHours(3)
        __Check(CStr(dtFuture.ToString("yyyy-MM-dd HH:mm")), "2025-01-06 13:00")
    End Sub
End Module
