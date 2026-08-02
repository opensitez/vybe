' vybe-test: vb/vb_sorted_list_operations/test_vb_sorted_list_auto_sorting
' origin: languages/vb/tests/vb/test_vb_sorted_list_operations.rs

Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim list As New SortedList(Of String, Integer)
        list.Add("Zebra", 26)
        list.Add("Apple", 1)
        list.Add("Monkey", 13)
        For Each kvp In list
            Console.WriteLine(kvp.Key & ":" & kvp.Value)
        Next
    End Sub
End Module
