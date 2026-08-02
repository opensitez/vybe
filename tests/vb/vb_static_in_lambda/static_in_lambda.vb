' vybe-test: vb/vb_static_in_lambda/static_in_lambda
' origin: languages/vb/tests/vb/test_vb_static_in_lambda.rs

Module M
    Sub Main()
        ' VB does not allow Static locals inside lambdas.
        ' This is purely to ensure the parser handles the error gracefully.
        ' We wrap it in a scenario that might parse if parser is permissive or correctly flags it.
        Dim act = Sub()
                      Static count As Integer = 0
                      count += 1
                      Console.WriteLine(count)
                  End Sub
                  
        act()
        act()
    End Sub
End Module
