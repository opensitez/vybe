' vybe-test: vb/vb_multidimensional_array_slicing/test_vb_array_2d_out_of_bounds_exception
' origin: languages/vb/tests/vb/test_vb_multidimensional_array_slicing.rs

Module Program
    Sub Main()
        Try
            Dim arr(1, 1) As Integer
            Dim x As Integer = arr(2, 0)
            Console.WriteLine(x)
        Catch ex As IndexOutOfRangeException
            Console.WriteLine("IndexOutOfRangeException")
        End Try
    End Sub
End Module
