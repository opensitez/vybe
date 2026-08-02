' vybe-test: vb/vb_floating_point_infinity_nan/test_vb_infinity_arithmetic_operations
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
        Dim inf = Double.PositiveInfinity
        Dim resAdd = inf + 100
        Dim resMult = inf * 2
        __Check(CStr(Double.IsPositiveInfinity(resAdd) & "|" & Double.IsPositiveInfinity(resMult)), "True|True")
    End Sub
End Module
