' vybe-test: vb/vb_full_domain_model_simulation/test_vb_domain_order_batch_processor_concurrent_queue
' origin: languages/vb/tests/vb/test_vb_full_domain_model_simulation.rs

Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim orderQueue As New ConcurrentQueue(Of String)()
        orderQueue.Enqueue("Ord1")
        orderQueue.Enqueue("Ord2")

        Dim processedCount = 0
        Dim id As String = Nothing
        While orderQueue.TryDequeue(id)
            processedCount += 1
        End While
        Console.WriteLine("Processed Orders: " & processedCount)
    End Sub
End Module
