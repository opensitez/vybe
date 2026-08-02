' vybe-test: vb/vb_property_expression_bodied/test_vb_function_single_line_lambda_expression
' origin: languages/vb/tests/vb/test_vb_property_expression_bodied.rs

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
    Public Function Multiply(x As Integer, y As Integer) As Integer => x * y
    Public Function IsPositive(n As Integer) As Boolean => n > 0

    Sub Main()
        __Check(CStr(Multiply(3, 4)), "12")
        __Check(CStr(IsPositive(5)), "True")
    End Sub
End Module
