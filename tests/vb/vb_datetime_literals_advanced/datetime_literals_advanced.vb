' vybe-test: vb/vb_datetime_literals_advanced/datetime_literals_advanced
' origin: languages/vb/tests/vb/test_vb_datetime_literals_advanced.rs

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
        ' ISO format YYYY-MM-DD
        Dim d1 As Date = #2024-05-15#
        
        ' With time YYYY-MM-DD HH:MM:SS
        Dim d2 As Date = #2024-05-15 14:30:00#
        
        ' AM/PM format
        Dim d3 As Date = #5/15/2024 2:30 PM#
        
        __Check(CStr(d1.Year), "2024")
        __Check(CStr(d2.Hour), "14")
        __Check(CStr(d3.Hour), "14")
    End Sub
End Module
