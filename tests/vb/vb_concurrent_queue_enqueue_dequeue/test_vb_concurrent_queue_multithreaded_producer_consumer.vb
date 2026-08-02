' vybe-test: vb/vb_concurrent_queue_enqueue_dequeue/test_vb_concurrent_queue_multithreaded_producer_consumer
' origin: languages/vb/tests/vb/test_vb_concurrent_queue_enqueue_dequeue.rs

Imports System.Collections.Concurrent
Imports System.Threading.Tasks

Module Program
    Sub Main()
        Dim q As New ConcurrentQueue(Of Integer)()
        Parallel.For(0, 50, Sub(i) q.Enqueue(i))

        Dim count = 0
        Dim val As Integer
        While q.TryDequeue(val)
            count += 1
        End While

        Console.WriteLine("Dequeued Total: " & count)
    End Sub
End Module
