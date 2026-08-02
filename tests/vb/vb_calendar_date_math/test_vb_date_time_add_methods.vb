' vybe-test: vb/vb_calendar_date_math/test_vb_date_time_add_methods
' origin: languages/vb/tests/vb/test_vb_calendar_date_math.rs

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
        Dim dt As New DateTime(2024, 2, 28) ' Leap year
        __Check(CStr(dt.AddDays(1).ToString("yyyy-MM-dd")), "2024-02-29")
        __Check(CStr(dt.AddMonths(1).ToString("yyyy-MM-dd")), "2024-03-28")
        __Check(CStr(dt.AddYears(1).ToString("yyyy-MM-dd")), "2025-02-28")
        __Check(CStr(DateTime.IsLeapYear(2024)), "True")
    End Sub
End Module
