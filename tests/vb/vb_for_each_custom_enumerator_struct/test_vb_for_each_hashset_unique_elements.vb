' vybe-test: vb/vb_for_each_custom_enumerator_struct/test_vb_for_each_hashset_unique_elements
' origin: languages/vb/tests/vb/test_vb_for_each_custom_enumerator_struct.rs

Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim set As New HashSet(Of Integer) From {1, 2, 2, 3}
        Console.WriteLine(set.Count)
        For Each item In set
            Console.WriteLine(item)
        Next
    End Sub
End Module
