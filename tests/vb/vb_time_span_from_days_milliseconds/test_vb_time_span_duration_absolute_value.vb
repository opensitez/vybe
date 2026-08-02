' vybe-test: vb/vb_time_span_from_days_milliseconds/test_vb_time_span_duration_absolute_value
' origin: languages/vb/tests/vb/test_vb_time_span_from_days_milliseconds.rs

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
        Dim neg = TimeSpan.FromHours(-5)
        Dim abs = neg.Duration()
        __Check(CStr(abs.TotalHours), "5")
    End Sub
End Module
