' vybe-test: vb/vb_date_literals_adv/date_literals_formats
' origin: languages/vb/tests/vb/test_vb_date_literals_adv.rs

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
        ' Multiple formats allowed in date literals
        Dim d1 As Date = #1998-11-23#
        Dim d2 As Date = #23 Nov 98#
        Dim d3 As Date = #1:15 PM#
        
        __Check(CStr(d1.Year), "1998")
        __Check(CStr(d2.Month), "11")
        __Check(CStr(d3.Hour), "13")
    End Sub
End Module
