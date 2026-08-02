' vybe-test: vb/vb_concurrent_queue_enqueue_dequeue/test_vb_concurrent_queue_linq_query_filtering
' origin: languages/vb/tests/vb/test_vb_concurrent_queue_enqueue_dequeue.rs

Imports System.Collections.Concurrent
Imports System.Linq

Module Program
    Sub Main()
        Dim q As New ConcurrentQueue(Of Integer)()
        For i As Integer = 1 To 10 : q.Enqueue(i) : Next
        Dim evens = q.Where(Function(n) n Mod 2 = 0).ToList()
        Console.WriteLine(String.Join(",", evens))
    End Sub
End Module
