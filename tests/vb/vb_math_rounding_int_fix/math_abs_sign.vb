' vybe-test: vb/vb_math_rounding_int_fix/math_abs_sign
' origin: languages/vb/tests/vb/test_vb_math_rounding_int_fix.rs

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
        __Check(CStr(Abs(-50.5)), "50.5")
        
        ' Sign returns -1, 0, or 1
        __Check(CStr(Sign(-100)), "-1")
        __Check(CStr(Sign(0)), "0")
        __Check(CStr(Sign(45)), "1")
    End Sub
End Module
