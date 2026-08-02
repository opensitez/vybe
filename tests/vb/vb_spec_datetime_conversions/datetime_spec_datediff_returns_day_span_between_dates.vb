' vybe-test: vb/vb_spec_datetime_conversions/datetime_spec_datediff_returns_day_span_between_dates
' origin: languages/vb/tests/vb/test_vb_spec_datetime_conversions.rs

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

Module M
    Sub Main()
        __Check(CStr(DateDiff(DateInterval.Day, #5/1/2024#, #5/14/2024#)), "13")
    End Sub
End Module
