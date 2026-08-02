' vybe-test: vb/vb_index_out_of_range_exception/test_vb_array_exceed_upper_bound_throws_index_out_of_range
' origin: languages/vb/tests/vb/test_vb_index_out_of_range_exception.rs

Imports System

Module Program
    Sub Main()
        Dim arr As Integer() = {10, 20, 30}
        Try
            Dim x As Integer = arr(3)
            Console.WriteLine(x)
        Catch ex As IndexOutOfRangeException
            Console.WriteLine("Exceeded Upper Bound Caught")
        End Try
    End Sub
End Module
