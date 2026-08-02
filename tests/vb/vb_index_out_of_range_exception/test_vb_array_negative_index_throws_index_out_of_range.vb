' vybe-test: vb/vb_index_out_of_range_exception/test_vb_array_negative_index_throws_index_out_of_range
' origin: languages/vb/tests/vb/test_vb_index_out_of_range_exception.rs

Imports System

Module Program
    Sub Main()
        Dim arr As Integer() = {10, 20, 30}
        Try
            Dim x As Integer = arr(-1)
            Console.WriteLine(x)
        Catch ex As IndexOutOfRangeException
            Console.WriteLine("Negative Index Caught")
        End Try
    End Sub
End Module
