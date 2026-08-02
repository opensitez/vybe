' vybe-test: vb/vb_datetime_comparisons/datetime_comparisons
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
        Dim d1 As Date = #2024-01-01 12:00:00#
        Dim d2 As Date = #2024-01-01 12:00:00#
        Dim d3 As Date = #2024-01-02#
        
        __Check(CStr(d1 = d2), "True")
        __Check(CStr(d1 < d3), "True")
        __Check(CStr(d3 >= d1), "True")
        
        ' Compare method
        __Check(CStr(Date.Compare(d1, d3)), "-1")
    End Sub
End Module
