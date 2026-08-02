' vybe-test: vb/vb_array_apis/array_copy_moves_values_between_arrays
' origin: languages/vb/tests/vb/test_vb_array_apis.rs

Module M
    Sub Main()
        Dim source As Integer() = {5, 6, 7}
        Dim target As Integer() = New Integer(2) {}
        Array.Copy(source, target, 3)
        For Each value As Integer In target
            Console.WriteLine(value)
        Next
    End Sub
End Module
