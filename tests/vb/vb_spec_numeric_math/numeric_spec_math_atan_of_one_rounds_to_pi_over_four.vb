' vybe-test: vb/vb_spec_numeric_math/numeric_spec_math_atan_of_one_rounds_to_pi_over_four
' origin: languages/vb/tests/vb/test_vb_spec_numeric_math.rs

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
        __Check(CStr(Round(Math.Atan(1), 6)), "0.785398")
    End Sub
End Module
