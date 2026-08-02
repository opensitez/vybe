' vybe-test: vb/vb_datetime_comparisons/datetime_parts
' origin: languages/vb/tests/vb/test_vb_datetime_comparisons.rs

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
        Dim d As Date = #2024-02-15 14:30:45#
        
        __Check(CStr(d.Year), "2024")
        __Check(CStr(d.Month), "2")
        __Check(CStr(d.Day), "15")
        __Check(CStr(d.Hour), "14")
        __Check(CStr(d.Minute), "30")
        __Check(CStr(d.Second), "45")
        __Check(CStr(d.DayOfWeek.ToString()), "Thursday")
    End Sub
End Module
