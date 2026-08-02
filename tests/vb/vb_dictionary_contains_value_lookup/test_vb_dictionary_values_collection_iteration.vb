' vybe-test: vb/vb_dictionary_contains_value_lookup/test_vb_dictionary_values_collection_iteration
' origin: languages/vb/tests/vb/test_vb_dictionary_contains_value_lookup.rs

Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New Dictionary(Of String, Integer) From {{"A", 10}, {"B", 20}}
        Dim sum As Integer = 0
        For Each val In dict.Values
            sum += val
        Next
        Console.WriteLine(sum)
    End Sub
End Module
