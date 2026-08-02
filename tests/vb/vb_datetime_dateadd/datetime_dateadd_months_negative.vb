' vybe-test: vb/vb_datetime_dateadd/datetime_dateadd_months_negative
' origin: languages/vb/tests/vb/test_vb_datetime_dateadd.rs

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
        Dim dt As Date = #5/15/2026#
        ' Subtract 2 months (DateInterval.Month)
        Dim newDt As Date = DateAdd("m", -2, dt)
        __Check(CStr(newDt.Month), "3")
        __Check(CStr(newDt.Year), "2026")
    End Sub
End Module
