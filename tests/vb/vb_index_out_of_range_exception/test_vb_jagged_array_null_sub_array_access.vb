' vybe-test: vb/vb_index_out_of_range_exception/test_vb_jagged_array_null_sub_array_access
' origin: languages/vb/tests/vb/test_vb_index_out_of_range_exception.rs

Imports System

Module Program
    Sub Main()
        Dim jagged As Integer()() = New Integer(2)() {}
        ' Sub-array at index 0 is Nothing
        Try
            Dim val As Integer = jagged(0)(0)
            Console.WriteLine(val)
        Catch ex As NullReferenceException
            Console.WriteLine("Null Sub-Array NullReferenceException Caught")
        End Try
    End Sub
End Module
