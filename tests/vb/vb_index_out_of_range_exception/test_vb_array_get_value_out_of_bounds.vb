' vybe-test: vb/vb_index_out_of_range_exception/test_vb_array_get_value_out_of_bounds
' origin: languages/vb/tests/vb/test_vb_index_out_of_range_exception.rs

Imports System

Module Program
    Sub Main()
        Dim arr As Array = New String() {"Alpha", "Beta"}
        Try
            Dim val As Object = arr.GetValue(10)
            Console.WriteLine(val)
        Catch ex As IndexOutOfRangeException
            Console.WriteLine("GetValue IndexOutOfRangeException Caught")
        End Try
    End Sub
End Module
