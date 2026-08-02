' vybe-test: vb/vb_date_time_compare_is_leap_year/test_vb_date_time_linq_max_min
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

Imports System.Linq

Module Program
    Sub Main()
        Dim dates = {New DateTime(2025, 12, 1), New DateTime(2025, 1, 1), New DateTime(2025, 6, 1)}
        __Check(CStr(dates.Min().Month & "|" & dates.Max().Month), "1|12")
    End Sub
End Module
