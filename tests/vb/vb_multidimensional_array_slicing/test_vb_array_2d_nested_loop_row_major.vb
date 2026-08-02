' vybe-test: vb/vb_multidimensional_array_slicing/test_vb_array_2d_nested_loop_row_major
' origin: languages/vb/tests/vb/test_vb_multidimensional_array_slicing.rs

Module Program
    Sub Main()
        Dim matrix(,) As Integer = {{1, 2}, {3, 4}}
        Dim sum As Integer = 0
        For i As Integer = 0 To matrix.GetUpperBound(0)
            For j As Integer = 0 To matrix.GetUpperBound(1)
                sum += matrix(i, j)
            Next
        Next
        Console.WriteLine(sum)
    End Sub
End Module
