' vybe-test: vb/vb_datetime_datediff/datetime_datediff_days
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
        Dim d1 As Date = #1/1/2026#
        Dim d2 As Date = #1/10/2026#
        ' Difference in days (d2 - d1)
        __Check(CStr(DateDiff(DateInterval.Day, d1, d2)), "9")
    End Sub
End Module
