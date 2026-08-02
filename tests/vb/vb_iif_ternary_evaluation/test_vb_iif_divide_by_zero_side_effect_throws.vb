' vybe-test: vb/vb_iif_ternary_evaluation/test_vb_iif_divide_by_zero_side_effect_throws
' origin: languages/vb/tests/vb/test_vb_iif_ternary_evaluation.rs

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
Imports Microsoft.VisualBasic

Module Program
    Sub Main()
        Try
            ' Because IIf evaluates both branches, 10 / 0 in falsepart throws DivideByZeroException even when condition is True!
            Dim res = IIf(True, 42, 10 \ 0)
        Catch ex As DivideByZeroException
            __Check(CStr("DivideByZeroException Caught in Eager IIf"), "DivideByZeroException Caught in Eager IIf")
        End Try
    End Sub
End Module
