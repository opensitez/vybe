' vybe-test: vb/vb_array_apis/array_copyto_moves_values_into_target
' origin: languages/vb/tests/vb/test_vb_array_apis.rs

Module M
    Sub Main()
        Dim source As Integer() = {9, 8}
        Dim target As Integer() = New Integer(1) {}
        source.CopyTo(target, 0)
        For Each value As Integer In target
            Console.WriteLine(value)
        Next
    End Sub
End Module
