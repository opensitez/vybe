' vybe-test: vb/vb_sorted_dictionary_custom_key_comparer/test_vb_sorted_dictionary_key_value_pair_iteration
' origin: languages/vb/tests/vb/test_vb_sorted_dictionary_custom_key_comparer.rs

Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of String, Integer) From {{"B", 2}, {"A", 1}}
        For Each kv In dict
            Console.WriteLine(kv.Key & "=" & kv.Value)
        Next
    End Sub
End Module
