' vybe-test: vb/vb_spec_delegates_lambdas/delegate_spec_lambda_inside_loop_can_use_iteration_value
' origin: languages/vb/tests/vb/test_vb_spec_delegates_lambdas.rs

Module M
    Sub Main()
        Dim fn As Func(Of Integer, Integer)
        For i As Integer = 1 To 3
            fn = Function(x) x + i
        Next
        Console.WriteLine(fn(4))
    End Sub
End Module
