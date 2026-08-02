' vybe-test: vb/vb_system_array_2d_matrix/array_2d_matrix_sum_rows_and_columns
' origin: languages/vb/tests/vb/test_vb_system_array_2d_matrix.rs

Module M
    Sub Main()
        Dim m(2, 1) As Integer
        m(0, 0) = 1
        m(0, 1) = 2
        m(1, 0) = 3
        m(1, 1) = 4
        m(2, 0) = 5
        m(2, 1) = 6

        Dim rowSums(2) As Integer
        For r As Integer = m.GetLowerBound(0) To m.GetUpperBound(0)
            For c As Integer = m.GetLowerBound(1) To m.GetUpperBound(1)
                rowSums(r) += m(r, c)
            Next
        Next

        Console.WriteLine(rowSums(0))
        Console.WriteLine(rowSums(1))
        Console.WriteLine(rowSums(2))

        Dim col0 As Integer = m(0, 0) + m(1, 0) + m(2, 0)
        Dim col1 As Integer = m(0, 1) + m(1, 1) + m(2, 1)
        Console.WriteLine(col0)
        Console.WriteLine(col1)
    End Sub
End Module
