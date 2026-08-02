' vybe-test: vb/vb_index_out_of_range_exception/test_vb_match_collection_regex_index_out_of_bounds
' origin: languages/vb/tests/vb/test_vb_index_out_of_range_exception.rs

Imports System
Imports System.Text.RegularExpressions

Module Program
    Sub Main()
        Dim matches = Regex.Matches("123 456", "\d+")
        Try
            Dim m = matches(10)
            Console.WriteLine(m.Value)
        Catch ex As ArgumentOutOfRangeException
            Console.WriteLine("MatchCollection ArgumentOutOfRangeException Caught")
        End Try
    End Sub
End Module
