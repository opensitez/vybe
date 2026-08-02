' vybe-test: vb/vb_for_each_custom_enumerator_struct/test_vb_for_each_stack_lifo_order
' origin: languages/vb/tests/vb/test_vb_for_each_custom_enumerator_struct.rs

Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim s As New Stack(Of String)()
        s.Push("First")
        s.Push("Second")
        For Each item In s
            Console.WriteLine(item)
        Next
    End Sub
End Module
