' vybe-test: vb/vb_math_trig_hyperbolic_exp/test_vb_math_trig_functions
' origin: languages/vb/tests/vb/test_vb_math_trig_hyperbolic_exp.rs

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
        __Check(CStr(Math.Sin(0.0)), "0")
        __Check(CStr(Math.Cos(0.0)), "1")
        __Check(CStr(Math.Log10(100.0)), "2")
        __Check(CStr(Math.Exp(0.0)), "1")
    End Sub
End Module
