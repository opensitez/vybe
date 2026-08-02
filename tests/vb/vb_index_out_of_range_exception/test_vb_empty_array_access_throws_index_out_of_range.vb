' vybe-test: vb/vb_index_out_of_range_exception/test_vb_empty_array_access_throws_index_out_of_range
' origin: languages/vb/tests/vb/test_vb_index_out_of_range_exception.rs

Imports System

Module Program
    Sub Main()
        Dim arr As Integer() = {}
        Try
            Dim x As Integer = arr(0)
            Console.WriteLine(x)
        Catch ex As IndexOutOfRangeException
            Console.WriteLine("Empty Array Access Caught")
        End Try
    End Sub
End Module
