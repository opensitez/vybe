' vybe-test: vb/vb_array_apis/array_convert_all_maps_to_strings
' origin: languages/vb/tests/vb/test_vb_array_apis.rs

Module M
    Sub Main()
        Dim values As Integer() = {1, 2, 3}
        Dim text As String() = Array.ConvertAll(values, Function(value As Integer) "n" & value)
        For Each part As String In text
            Console.WriteLine(part)
        Next
    End Sub
End Module
