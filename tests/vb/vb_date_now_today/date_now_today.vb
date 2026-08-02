' vybe-test: vb/vb_date_now_today/date_now_today
' origin: languages/vb/tests/vb/test_vb_date_now_today.rs

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
        ' We can't easily assert the exact time, but we can verify types and properties
        Dim n As Date = Now
        Dim t As Date = Today
        
        __Check(CStr(n.Year >= 2020), "True")
        __Check(CStr(t.TimeOfDay.TotalSeconds = 0), "True") ' Today has no time component (midnight)
        
        Dim tod As Date = TimeOfDay
        __Check(CStr(tod.Year = 1), "True") ' TimeOfDay has dummy date component
    End Sub
End Module
