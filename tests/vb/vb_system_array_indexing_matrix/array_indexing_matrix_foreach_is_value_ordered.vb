' vybe-test: vb/vb_system_array_indexing_matrix/array_indexing_matrix_foreach_is_value_ordered
' origin: languages/vb/tests/vb/test_vb_system_array_indexing_matrix.rs

Module M
    Sub Main()
        Dim values As Integer() = {4, 1, 7, 0}
        Dim ordered As New System.Text.StringBuilder()

        For Each value As Integer In values
            ordered.Append(value).Append(",")
        Next

        Console.WriteLine(ordered.ToString())
    End Sub
End Module
