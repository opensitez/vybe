' vybe-test: vb/vb_lambda_scopes/lambda_closures_loop_variable
' origin: languages/vb/tests/vb/test_vb_lambda_scopes.rs

Module M
    Sub Main()
        Dim funcs As New System.Collections.Generic.List(Of Func(Of Integer))
        
        ' In VB.NET, capturing the loop variable directly inside a For loop captures the same variable
        ' so it will evaluate to the final value (4).
        For i As Integer = 1 To 3
            funcs.Add(Function() i * 2)
        Next
        
        For Each f In funcs
            Console.WriteLine(f())
        Next
    End Sub
End Module
