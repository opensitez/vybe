' vybe-test: vb/vb_datetime_datediff/datetime_datediff_hours_negative
' origin: languages/vb/tests/vb/test_vb_datetime_datediff.rs

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
        Dim d1 As Date = #1/2/2026 12:00:00 PM#
        Dim d2 As Date = #1/2/2026 8:00:00 AM#
        ' Difference in hours (d2 - d1), should be negative
        __Check(CStr(DateDiff("h", d1, d2)), "-4")
    End Sub
End Module
