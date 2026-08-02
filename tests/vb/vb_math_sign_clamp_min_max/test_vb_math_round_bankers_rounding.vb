' vybe-test: vb/vb_math_sign_clamp_min_max/test_vb_math_round_bankers_rounding
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
        ' Banker's Rounding (Round to Even): 2.5 -> 2, 3.5 -> 4
        __Check(CStr(Math.Round(2.5) & "|" & Math.Round(3.5)), "2|4")
    End Sub
End Module
