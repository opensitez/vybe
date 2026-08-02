' vybe-test: vb/vb_tuple_hashset_dictionary_keys/test_vb_tuple_sorted_set_ordering
' origin: languages/vb/tests/vb/test_vb_tuple_hashset_dictionary_keys.rs

Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim set As New SortedSet(Of (Integer, String)) From {
            (2, "B"),
            (1, "Z"),
            (1, "A")
        }
        For Each item In set
            Console.WriteLine(item.Item1 & ":" & item.Item2)
        Next
    End Sub
End Module
