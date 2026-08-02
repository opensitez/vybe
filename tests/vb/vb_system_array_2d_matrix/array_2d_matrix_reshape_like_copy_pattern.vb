' vybe-test: vb/vb_system_array_2d_matrix/array_2d_matrix_reshape_like_copy_pattern
' origin: languages/vb/tests/vb/test_vb_system_array_2d_matrix.rs

Module M
    Sub Main()
        Dim source(1, 2) As Integer
        Dim value As Integer = 1
        For i As Integer = source.GetLowerBound(0) To source.GetUpperBound(0)
            For j As Integer = source.GetLowerBound(1) To source.GetUpperBound(1)
                source(i, j) = value
                value += 1
            Next
        Next

        Dim copied(1, 2) As Integer
        Array.Copy(source, copied, source.Length)

        Console.WriteLine(copied(0, 0))
        Console.WriteLine(copied(0, 2))
        Console.WriteLine(copied(1, 1))
    End Sub
End Module
