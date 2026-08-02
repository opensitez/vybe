' vybe-test: vb/vb_index_out_of_range_exception/test_vb_list_out_of_bounds_throws_argument_out_of_range
' origin: languages/vb/tests/vb/test_vb_index_out_of_range_exception.rs

Imports System
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim list As New List(Of Integer) From {1, 2}
        Try
            Dim x As Integer = list(5)
            Console.WriteLine(x)
        Catch ex As ArgumentOutOfRangeException
            Console.WriteLine("List ArgumentOutOfRangeException Caught")
        End Try
    End Sub
End Module
