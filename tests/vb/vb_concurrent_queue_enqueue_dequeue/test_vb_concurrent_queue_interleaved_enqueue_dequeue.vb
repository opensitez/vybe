' vybe-test: vb/vb_concurrent_queue_enqueue_dequeue/test_vb_concurrent_queue_interleaved_enqueue_dequeue
' origin: languages/vb/tests/vb/test_vb_concurrent_queue_enqueue_dequeue.rs

' Vybe test harness — Visual Basic.
'
' Real VB source alongside harness/go/check.go and harness/js/check.js, the way
' test262's assert.js is JavaScript.
'
' A test's verdict is its EXIT CODE. __Check prints its diagnostic BEFORE
' throwing: an uncaught exception surfaces as `RuntimeError: [object]`, which
' says nothing at all.

Module VybeCheck
    Sub __Check(got As String, want As String)
        If got <> want Then
            Console.WriteLine("FAIL: want [" & want & "] got [" & got & "]")
            Throw New Exception("assertion failed")
        End If
    End Sub
End Module

Imports System.Collections.Concurrent

Module Program
    Sub Main()
        Dim q As New ConcurrentQueue(Of Integer)()
        q.Enqueue(1)
        q.Enqueue(2)
        Dim v1 As Integer
        q.TryDequeue(v1)

        q.Enqueue(3)
        Dim v2 As Integer, v3 As Integer
        q.TryDequeue(v2)
        q.TryDequeue(v3)

        __Check(CStr(v1 & "|" & v2 & "|" & v3), "1|2|3")
    End Sub
End Module
