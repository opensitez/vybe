' vybe-test: vb/vb_array_apis/array_clear_resets_range_to_default
' origin: languages/vb/tests/vb/test_vb_array_apis.rs

Module M
    Sub Main()
        Dim values As Integer() = {1, 2, 3}
        Array.Clear(values, 1, 2)
        For Each value As Integer In values
            Console.WriteLine(value)
        Next
    End Sub
End Module
