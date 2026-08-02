' vybe-test: vb/vb_concurrent_queue_enqueue_dequeue/test_vb_concurrent_queue_multiple_dequeues_in_parallel
' origin: languages/vb/tests/vb/test_vb_concurrent_queue_enqueue_dequeue.rs

Imports System.Collections.Concurrent
Imports System.Threading.Tasks

Module Program
    Sub Main()
        Dim q As New ConcurrentQueue(Of Integer)()
        For i As Integer = 1 To 100 : q.Enqueue(i) : Next

        Dim sum = 0
        Dim lockObj As New Object()
        Parallel.For(0, 100, Sub(i)
            Dim item As Integer
            If q.TryDequeue(item) Then
                SyncLock lockObj
                    sum += item
                End SyncLock
            End If
        End Sub)
        Console.WriteLine(sum & "|QueueEmpty=" & q.IsEmpty)
    End Sub
End Module
