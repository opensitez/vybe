' vybe-test: vb/vb_time_span_arithmetic_units/test_vb_time_span_total_units
' origin: languages/vb/tests/vb/test_vb_time_span_arithmetic_units.rs

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
        Dim ts As TimeSpan = TimeSpan.FromHours(2.5)
        __Check(CStr(ts.TotalMinutes), "150")
        __Check(CStr(ts.TotalSeconds), "9000")
        __Check(CStr(ts.Hours & ":" & ts.Minutes), "2:30")
    End Sub
End Module
