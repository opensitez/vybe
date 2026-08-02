' vybe-test: vb/vb_date_arithmetic/date_arithmetic
' origin: languages/vb/tests/vb/test_vb_date_arithmetic.rs

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
        Dim d1 As Date = #1/1/2024#
        Dim d2 As Date = #1/5/2024#
        
        ' Subtracting dates returns a TimeSpan
        Dim ts As TimeSpan = d2 - d1
        __Check(CStr(ts.Days), "4")
        
        ' Adding TimeSpan to Date
        Dim d3 As Date = d1 + New TimeSpan(10, 0, 0, 0)
        __Check(CStr(d3.Day), "11")
    End Sub
End Module
