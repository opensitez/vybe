' vybe-test: vb/vb_tuple_hashset_dictionary_keys/test_vb_tuple_sorted_dictionary_custom_key_ordering
' origin: languages/vb/tests/vb/test_vb_tuple_hashset_dictionary_keys.rs

Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of (Integer, Integer), String)()
        dict((2, 1)) = "P21"
        dict((1, 5)) = "P15"
        dict((1, 2)) = "P12"

        For Each kv In dict
            Console.WriteLine(kv.Key.Item1 & "," & kv.Key.Item2 & "=" & kv.Value)
        Next
    End Sub
End Module
