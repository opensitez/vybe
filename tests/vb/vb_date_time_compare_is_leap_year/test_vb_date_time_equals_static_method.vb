' vybe-test: vb/vb_date_time_compare_is_leap_year/test_vb_date_time_equals_static_method
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
    Sub Main()
        Dim d1 As New DateTime(2025, 1, 1, 10, 0, 0)
        Dim d2 As New DateTime(2025, 1, 1, 10, 0, 0)
        Dim d3 As New DateTime(2025, 1, 1, 11, 0, 0)
        __Check(CStr(DateTime.Equals(d1, d2) & "|" & DateTime.Equals(d1, d3)), "True|False")
    End Sub
End Module
