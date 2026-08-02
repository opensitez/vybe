' vybe-test: vb/vb_math_sign_clamp_min_max/test_vb_math_scaleb_floating_point
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
        ' ScaleB(x, n) calculates x * 2^n
        Dim res = Math.ScaleB(1.5, 3) ' 1.5 * 8 = 12
        __Check(CStr(res), "12")
    End Sub
End Module
