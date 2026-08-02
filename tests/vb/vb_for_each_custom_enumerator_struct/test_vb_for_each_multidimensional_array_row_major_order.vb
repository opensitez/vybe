' vybe-test: vb/vb_for_each_custom_enumerator_struct/test_vb_for_each_multidimensional_array_row_major_order
' origin: languages/vb/tests/vb/test_vb_for_each_custom_enumerator_struct.rs

Module Program
    Sub Main()
        Dim grid(,) As Integer = {{1, 2}, {3, 4}}
        Dim res = ""
        For Each val In grid
            res &= val & ","
        Next
        Console.WriteLine(res.TrimEnd(","c))
    End Sub
End Module
