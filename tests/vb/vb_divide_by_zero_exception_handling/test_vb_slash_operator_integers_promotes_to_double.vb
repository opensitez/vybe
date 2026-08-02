' vybe-test: vb/vb_divide_by_zero_exception_handling/test_vb_slash_operator_integers_promotes_to_double
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
        Dim a As Integer = 10
        Dim b As Integer = 0
        ' Visual Basic "/" operator between integers performs floating point division!
        Dim res As Double = a / b
        __Check(CStr(Double.IsInfinity(res)), "True")
    End Sub
End Module
