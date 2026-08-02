' vybe-test: vb/vb_math_sign_round/math_sign_round
' origin: languages/vb/tests/vb/test_vb_math_sign_round.rs

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
        ' Sign returns 1, 0, or -1
        __Check(CStr(Sign(-42)), "-1")
        __Check(CStr(Sign(0)), "0")
        __Check(CStr(Sign(42)), "1")
        
        ' Round performs banker's rounding by default
        __Check(CStr(Round(2.5)), "2") ' Rounds to nearest even -> 2
        __Check(CStr(Round(3.5)), "4") ' Rounds to nearest even -> 4
    End Sub
End Module
