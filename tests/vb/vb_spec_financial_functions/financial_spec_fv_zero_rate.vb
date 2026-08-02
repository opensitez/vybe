' vybe-test: vb/vb_spec_financial_functions/financial_spec_fv_zero_rate
' origin: languages/vb/tests/vb/test_vb_spec_financial_functions.rs

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
        __Check(CStr(Round(FV(0, 10, -100, 0, 0), 2)), "1000")
    End Sub
End Module
