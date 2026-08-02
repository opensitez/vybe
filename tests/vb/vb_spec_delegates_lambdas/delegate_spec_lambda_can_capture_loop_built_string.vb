' vybe-test: vb/vb_spec_delegates_lambdas/delegate_spec_lambda_can_capture_loop_built_string
' origin: languages/vb/tests/vb/test_vb_spec_delegates_lambdas.rs

Module M
    Sub Main()
        Dim text As String = ""
        For i As Integer = 1 To 3
            text &= i
        Next
        Dim fn As Func(Of String) = Function() text
        Console.WriteLine(fn())
    End Sub
End Module
