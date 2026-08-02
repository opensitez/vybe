' vybe-test: vb/vb_lambdas/multiline_function_lambda
' origin: languages/vb/tests/vb/test_vb_lambdas.rs

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

Module M
    Sub Main()
        Dim factorial As Func(Of Integer, Integer) = Function(n)
            If n <= 1 Then Return 1
            Return n * factorial(n - 1)
        End Function
        __Check(CStr(factorial(5)), "120")
    End Sub
End Module
