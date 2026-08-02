' vybe-test: vb/vb_index_out_of_range_exception/test_vb_string_substring_index_out_of_bounds_throws_argument_out_of_range
' origin: languages/vb/tests/vb/test_vb_index_out_of_range_exception.rs

Imports System

Module Program
    Sub Main()
        Dim str As String = "ABC"
        Try
            Dim subStr As String = str.Substring(0, 10)
            Console.WriteLine(subStr)
        Catch ex As ArgumentOutOfRangeException
            Console.WriteLine("Substring ArgumentOutOfRangeException Caught")
        End Try
    End Sub
End Module
