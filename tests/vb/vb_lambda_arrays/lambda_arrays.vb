' vybe-test: vb/vb_lambda_arrays/lambda_arrays
' origin: languages/vb/tests/vb/test_vb_lambda_arrays.rs

Module M
    Sub Main()
        ' Lambda returning an array literal
        Dim getArray = Function() {1, 2, 3}
        
        Dim arr = getArray()
        For Each n In arr
            Console.WriteLine(n)
        Next
    End Sub
End Module
