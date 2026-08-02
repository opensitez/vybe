' vybe-test: vb/vb_floating_point_infinity_nan/test_vb_double_negative_zero_detection
' origin: languages/vb/tests/vb/test_vb_floating_point_infinity_nan.rs

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
        Dim posZero = 0.0
        Dim negZero = -0.0
        __Check(CStr((posZero = negZero) & "|" & (1.0 / posZero > 0) & "|" & (1.0 / negZero < 0)), "True|True|True")
    End Sub
End Module
