' vybe-test: vb/vb_linq_distinct_custom_equality_comparer/test_vb_linq_distinct_datetime_dates_only
' origin: languages/vb/tests/vb/test_vb_linq_distinct_custom_equality_comparer.rs

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
Imports System.Linq

Module Program
    Sub Main()
        Dim dates = {New DateTime(2025, 1, 1, 10, 0, 0), New DateTime(2025, 1, 1, 14, 0, 0), New DateTime(2025, 1, 2, 9, 0, 0)}
        Dim uniqueDays = dates.DistinctBy(Function(d) d.Date)
        __Check(CStr(uniqueDays.Count()), "2")
    End Sub
End Module
