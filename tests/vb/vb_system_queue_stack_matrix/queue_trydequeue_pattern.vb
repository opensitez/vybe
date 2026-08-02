' vybe-test: vb/vb_system_queue_stack_matrix/queue_trydequeue_pattern
' origin: languages/vb/tests/vb/test_vb_system_queue_stack_matrix.rs

Imports System.Collections.Generic

Module M
    Sub Main()
        Dim queue As New Queue(Of String)()
        queue.Enqueue("x")
        Dim first As String = queue.Dequeue()
        Console.WriteLine(first)
        Try
            queue.Dequeue()
            Console.WriteLine("extra")
        Catch ex As InvalidOperationException
            Console.WriteLine("empty")
        End Try
    End Sub
End Module
