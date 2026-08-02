' vybe-test: vb/vb_for_each_custom_enumerator_struct/test_vb_for_each_linked_list_traversal
' origin: languages/vb/tests/vb/test_vb_for_each_custom_enumerator_struct.rs

Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim list As New LinkedList(Of String)()
        list.AddLast("Node1")
        list.AddLast("Node2")
        For Each node In list
            Console.WriteLine(node)
        Next
    End Sub
End Module
