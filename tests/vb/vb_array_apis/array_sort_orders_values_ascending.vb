' vybe-test: vb/vb_array_apis/array_sort_orders_values_ascending
' origin: languages/vb/tests/vb/test_vb_array_apis.rs

Module M
    Sub Main()
        Dim values As Integer() = {4, 1, 3}
        Array.Sort(values)
        For Each value As Integer In values
            Console.WriteLine(value)
        Next
    End Sub
End Module
