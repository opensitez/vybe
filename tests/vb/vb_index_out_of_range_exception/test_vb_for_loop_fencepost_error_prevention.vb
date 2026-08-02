' vybe-test: vb/vb_index_out_of_range_exception/test_vb_for_loop_fencepost_error_prevention
' origin: languages/vb/tests/vb/test_vb_index_out_of_range_exception.rs

Module Program
    Sub Main()
        Dim arr As String() = {"A", "B", "C"}
        ' Valid 0 To arr.Length - 1
        For i As Integer = 0 To arr.Length - 1
            Console.WriteLine(arr(i))
        Next
    End Sub
End Module
