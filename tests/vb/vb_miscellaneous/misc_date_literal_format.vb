' vybe-test: vb/vb_miscellaneous/misc_date_literal_format
' origin: languages/vb/tests/vb/test_vb_miscellaneous.rs

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

Module M: Sub Main(): Dim d = #8/24/2020 12:30:00 PM#: __Check(CStr(d.Year), "2020"): End Sub: End Module
