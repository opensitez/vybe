' vybe-test: vb/vb_tuple_comparison_sorting/test_vb_tuple_list_sort_lexicographical
' origin: languages/vb/tests/vb/test_vb_tuple_comparison_sorting.rs

Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim list As New List(Of (Integer, String)) From {
            (2, "B"),
            (1, "Z"),
            (1, "A")
        }
        list.Sort()

        For Each item In list
            Console.WriteLine(item.Item1 & ":" & item.Item2)
        Next
    End Sub
End Module
