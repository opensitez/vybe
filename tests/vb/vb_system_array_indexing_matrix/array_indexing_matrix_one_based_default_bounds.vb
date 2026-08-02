' vybe-test: vb/vb_system_array_indexing_matrix/array_indexing_matrix_one_based_default_bounds
' origin: languages/vb/tests/vb/test_vb_system_array_indexing_matrix.rs

Module M
    Sub Main()
        Dim values(1 To 5) As Integer
        For i As Integer = 1 To 5
            values(i) = i * i
        Next

        Console.WriteLine(values.Length)
        Console.WriteLine(values(1))
        Console.WriteLine(values(5))
    End Sub
End Module
