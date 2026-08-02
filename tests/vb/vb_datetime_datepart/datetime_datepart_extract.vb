' vybe-test: vb/vb_datetime_datepart/datetime_datepart_extract
' origin: languages/vb/tests/vb/test_vb_datetime_datepart.rs

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
        Dim dt As Date = #12/25/2026 14:30:45#
        __Check(CStr(DatePart(DateInterval.Year, dt)), "2026")
        __Check(CStr(DatePart(DateInterval.Month, dt)), "12")
        __Check(CStr(DatePart(DateInterval.Day, dt)), "25")
        __Check(CStr(DatePart("h", dt)), "14")
        __Check(CStr(DatePart("n", dt)), "30") ' Minutes
        __Check(CStr(DatePart("s", dt)), "45") ' Seconds
    End Sub
End Module
