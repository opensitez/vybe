' vybe-test: vb/vb_floating_point_special_values/test_vb_float_division_by_zero_double
' origin: languages/vb/tests/vb/test_vb_floating_point_special_values.rs

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

Module Program
    Sub Main()
        Dim res1 As Double = 1.0 / 0.0
        Dim res2 As Double = -1.0 / 0.0
        Dim res3 As Double = 0.0 / 0.0
        __Check(CStr(Double.IsPositiveInfinity(res1)), "True")
        __Check(CStr(Double.IsNegativeInfinity(res2)), "True")
        __Check(CStr(Double.IsNaN(res3)), "True")
    End Sub
End Module
