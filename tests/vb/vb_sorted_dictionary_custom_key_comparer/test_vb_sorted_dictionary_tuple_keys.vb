' vybe-test: vb/vb_sorted_dictionary_custom_key_comparer/test_vb_sorted_dictionary_tuple_keys
' origin: languages/vb/tests/vb/test_vb_sorted_dictionary_custom_key_comparer.rs

Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of (Integer, Integer), String)()
        dict((2, 1)) = "B"
        dict((1, 5)) = "A"
        For Each kv In dict
            Console.WriteLine(kv.Key.Item1 & "," & kv.Key.Item2 & ":" & kv.Value)
        Next
    End Sub
End Module
