' vybe-test: vb/vb_date_time_compare_is_leap_year/test_vb_date_time_between_range_check
' origin: languages/vb/tests/vb/test_vb_date_time_compare_is_leap_year.rs

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
    Private Function IsInRange(target As DateTime, startDt As DateTime, endDt As DateTime) As Boolean
        Return target >= startDt AndAlso target <= endDt
    End Function

    Sub Main()
        Dim startDt As New DateTime(2025, 1, 1)
        Dim endDt As New DateTime(2025, 12, 31)
        Dim testDt As New DateTime(2025, 6, 15)
        __Check(CStr(IsInRange(testDt, startDt, endDt)), "True")
    End Sub
End Module
