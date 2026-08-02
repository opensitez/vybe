' vybe-test: vb/vb_index_out_of_range_exception/test_vb_jagged_array_sub_array_index_out_of_bounds
' origin: languages/vb/tests/vb/test_vb_index_out_of_range_exception.rs

Imports System

Module Program
    Sub Main()
        Dim jagged As Integer()() = New Integer(1)() {}
        jagged(0) = New Integer() {10, 20}
        Try
            Dim val As Integer = jagged(0)(5)
            Console.WriteLine(val)
        Catch ex As IndexOutOfRangeException
            Console.WriteLine("Jagged Sub-Array IndexOutOfRangeException Caught")
        End Try
    End Sub
End Module
