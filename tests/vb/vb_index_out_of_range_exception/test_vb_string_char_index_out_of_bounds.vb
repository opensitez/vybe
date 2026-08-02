' vybe-test: vb/vb_index_out_of_range_exception/test_vb_string_char_index_out_of_bounds
' origin: languages/vb/tests/vb/test_vb_index_out_of_range_exception.rs

Imports System

Module Program
    Sub Main()
        Dim text As String = "Hello"
        Try
            Dim ch As Char = text(10)
            Console.WriteLine(ch)
        Catch ex As IndexOutOfRangeException
            Console.WriteLine("String Index Out Of Range Caught")
        End Try
    End Sub
End Module
