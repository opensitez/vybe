' vybe-test: vb/vb_index_out_of_range_exception/test_vb_for_loop_upper_bound_syntax
' origin: languages/vb/tests/vb/test_vb_index_out_of_range_exception.rs

Module Program
    Sub Main()
        Dim arr As String() = {"X", "Y"}
        ' In VB, UBound(arr) equals arr.Length - 1
        For i As Integer = 0 To UBound(arr)
            Console.WriteLine(arr(i))
        Next
    End Sub
End Module
