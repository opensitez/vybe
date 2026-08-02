' vybe-test: vb/vb_array_redim_preserve/array_redim_no_preserve
' origin: languages/vb/tests/vb/test_vb_array_redim_preserve.rs

Module M
    Sub Main()
        Dim arr() As Integer = {1, 2, 3}
        
        ' ReDim without Preserve clears elements (initializes to default)
        ReDim arr(2)
        arr(0) = 9
        
        For i As Integer = 0 To UBound(arr)
            Console.WriteLine(arr(i))
        Next
    End Sub
End Module
