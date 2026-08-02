' vybe-test: vb/vb_concurrent_queue_enqueue_dequeue/test_vb_concurrent_queue_enumeration_snapshot_semantic
' origin: languages/vb/tests/vb/test_vb_concurrent_queue_enqueue_dequeue.rs

Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim q As New ConcurrentQueue(Of Integer)()
        q.Enqueue(1)
        q.Enqueue(2)

        Dim res = ""
        For Each item In q
            res &= item & ","
            If item = 1 Then q.Enqueue(3) ' Mutation does not affect active enumerator snapshot!
        Next
        Console.WriteLine(res.TrimEnd(","c) & "|Count=" & q.Count)
    End Sub
End Module
