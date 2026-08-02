' vybe-test: vb/vb_divide_by_zero_exception_handling/test_vb_single_float_zero_by_zero_returns_nan
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
        Dim a As Single = 0.0F
        Dim b As Single = 0.0F
        Dim res As Single = a / b
        __Check(CStr(Single.IsNaN(res)), "True")
    End Sub
End Module
