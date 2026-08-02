' vybe-test: vb/vb_named_tuple_nested_structures/test_vb_named_tuple_list_sorting_by_tuple_element
' origin: languages/vb/tests/vb/test_vb_named_tuple_nested_structures.rs

Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim items As New List(Of (Name As String, Priority As Integer)) From {
            ("TaskB", 2),
            ("TaskA", 1),
            ("TaskC", 3)
        }
        items.Sort(Function(x, y) x.Priority.CompareTo(y.Priority))
        For Each item In items
            Console.WriteLine(item.Name & ":" & item.Priority)
        Next
    End Sub
End Module
