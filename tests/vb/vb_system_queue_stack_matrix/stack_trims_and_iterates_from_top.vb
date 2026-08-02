' vybe-test: vb/vb_system_queue_stack_matrix/stack_trims_and_iterates_from_top
' origin: languages/vb/tests/vb/test_vb_system_queue_stack_matrix.rs

Imports System.Collections.Generic

Module M
    Sub Main()
        Dim stack As New Stack(Of String)()
        stack.Push("first")
        stack.Push("second")
        stack.Push("third")
        Dim output As String = ""
        For Each value As String In stack
            output &= value & ","
        Next
        Console.WriteLine(output)
        Console.WriteLine(stack.Count = 3)
    End Sub
End Module
