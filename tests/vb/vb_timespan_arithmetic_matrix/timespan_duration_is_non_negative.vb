' vybe-test: vb/vb_timespan_arithmetic_matrix/timespan_duration_is_non_negative
' origin: languages/vb/tests/vb/test_vb_timespan_arithmetic_matrix.rs

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
        Dim negative As TimeSpan = TimeSpan.FromMinutes(-5)
        __Check(CStr(negative.Duration().TotalMinutes), "5")
    End Sub
End Module
