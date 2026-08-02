' vybe-test: vb/vb_array_apis/array_resize_grows_array_and_preserves_existing_values
' origin: languages/vb/tests/vb/test_vb_array_apis.rs

Module M
    Sub Main()
        Dim values As Integer() = {2, 4}
        Array.Resize(values, 4)
        For Each value As Integer In values
            Console.WriteLine(value)
        Next
    End Sub
End Module
