' vybe-test: vb/vb_spec_delegates_lambdas/delegate_spec_lambda_can_return_another_lambda
' origin: languages/vb/tests/vb/test_vb_spec_delegates_lambdas.rs

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
    Function MakeFactory() As Func(Of Func(Of Integer, Integer))
        Return Function() Function(x) x * 3
    End Function
    Sub Main()
        __Check(CStr(MakeFactory()()(4)), "12")
    End Sub
End Module
