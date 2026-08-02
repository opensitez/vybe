' vybe-test: vb/vb_multidimensional_array_slicing/test_vb_array_2d_fill_diagonal
' origin: languages/vb/tests/vb/test_vb_multidimensional_array_slicing.rs

Module Program
    Sub Main()
        Dim identity(2, 2) As Integer
        For i As Integer = 0 To 2
            identity(i, i) = 1
        Next
        Console.WriteLine(identity(0, 0))
        Console.WriteLine(identity(0, 1))
        Console.WriteLine(identity(1, 1))
    End Sub
End Module
