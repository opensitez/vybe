' vybe-test: vb/vb_divide_by_zero_exception_handling/test_vb_infinity_multiplication_with_zero_is_nan
' origin: languages/vb/tests/vb/test_vb_divide_by_zero_exception_handling.rs

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
        Dim inf As Double = Double.PositiveInfinity
        Dim zero As Double = 0.0
        Dim result As Double = inf * zero
        __Check(CStr(Double.IsNaN(result)), "True")
    End Sub
End Module
