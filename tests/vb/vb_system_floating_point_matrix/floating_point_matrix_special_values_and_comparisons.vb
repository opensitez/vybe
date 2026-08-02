' vybe-test: vb/vb_system_floating_point_matrix/floating_point_matrix_special_values_and_comparisons
' origin: languages/vb/tests/vb/test_vb_system_floating_point_matrix.rs

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

Module M
    Sub Main()
        Dim n As Double = Double.NaN
        Dim p As Double = Double.PositiveInfinity
        Dim z As Double = Double.NegativeInfinity

        __Check(CStr(Double.IsNaN(n)), "True")
        __Check(CStr(Double.IsInfinity(p)), "True")
        __Check(CStr(Double.IsInfinity(z)), "True")
        __Check(CStr(n = n), "False")
        __Check(CStr(p > z), "True")
    End Sub
End Module
