' vybe-test: vb/vb_lambda_scopes/lambda_closures
' origin: languages/vb/tests/vb/test_vb_lambda_scopes.rs

Module M
    Sub Main()
        Dim funcs As New System.Collections.Generic.List(Of Func(Of Integer))
        
        ' VB.NET For loop variable is captured per iteration in modern versions (like C#)
        ' Actually, in VB.NET, the loop variable is declared outside the loop if not explicitly scoped
        ' Let's declare it inside the loop to be safe and test capture semantics.
        For i As Integer = 1 To 3
            Dim captured = i
            funcs.Add(Function() captured * 2)
        Next
        
        For Each f In funcs
            Console.WriteLine(f())
        Next
    End Sub
End Module
