' vybe-test: vb/vb_date_literals/date_literal_can_include_time_of_day_spec
' origin: languages/vb/tests/vb/test_vb_date_literals.rs

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
        Dim d As Date = #5/14/2024 3:45 PM#
        __Check(CStr(CStr(d)), "5/14/2024 3:45 PM")
    End Sub
End Module
