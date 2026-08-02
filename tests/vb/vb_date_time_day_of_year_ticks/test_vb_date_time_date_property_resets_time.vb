' vybe-test: vb/vb_date_time_day_of_year_ticks/test_vb_date_time_date_property_resets_time
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
        Dim dt As New DateTime(2025, 5, 20, 18, 45, 12)
        Dim dOnly = dt.Date
        __Check(CStr(dOnly.Hour & ":" & dOnly.Minute & ":" & dOnly.Second), "0:0:0")
    End Sub
End Module
