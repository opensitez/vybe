' vybe-test: vb/vb_date_add_diff/date_add_diff
' origin: languages/vb/tests/vb/test_vb_date_add_diff.rs

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
        Dim d1 As Date = #1/1/2020#
        
        ' DateAdd
        Dim d2 = DateAdd(DateInterval.Day, 10, d1)
        __Check(CStr(d2.Day), "11")
        
        ' DateDiff
        Dim diff = DateDiff(DateInterval.Day, d1, d2)
        __Check(CStr(diff), "10")
    End Sub
End Module
