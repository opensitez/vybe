' vybe-test: vb/vb_math_sign_clamp_min_max/test_vb_math_clamp_within_below_above_range
' origin: languages/vb/tests/vb/test_vb_math_sign_clamp_min_max.rs

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

Imports System

Module Program
    Sub Main()
        __Check(CStr(Math.Clamp(50, 0, 100) & "|" & Math.Clamp(-10, 0, 100) & "|" & Math.Clamp(150, 0, 100)), "50|0|100")
    End Sub
End Module
