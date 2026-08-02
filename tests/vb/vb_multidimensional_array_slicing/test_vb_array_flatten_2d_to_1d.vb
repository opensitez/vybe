' vybe-test: vb/vb_multidimensional_array_slicing/test_vb_array_flatten_2d_to_1d
' origin: languages/vb/tests/vb/test_vb_multidimensional_array_slicing.rs

Module Program
    Sub Main()
        Dim grid(,) As Integer = {{1, 2}, {3, 4}}
        Dim flat(grid.Length - 1) As Integer
        Dim idx As Integer = 0
        For Each val In grid
            flat(idx) = val
            idx += 1
        Next
        Console.WriteLine(String.Join(",", flat))
    End Sub
End Module
