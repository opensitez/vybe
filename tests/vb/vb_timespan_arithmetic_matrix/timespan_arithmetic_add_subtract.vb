' vybe-test: vb/vb_timespan_arithmetic_matrix/timespan_arithmetic_add_subtract
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
        Dim t1 As TimeSpan = TimeSpan.FromHours(1)
        Dim t2 As TimeSpan = TimeSpan.FromMinutes(30)
        Dim sum As TimeSpan = t1 + t2
        Dim diff As TimeSpan = t1 - t2
        __Check(CStr(sum.TotalMinutes), "90")
        __Check(CStr(diff.TotalMinutes), "30")
    End Sub
End Module
