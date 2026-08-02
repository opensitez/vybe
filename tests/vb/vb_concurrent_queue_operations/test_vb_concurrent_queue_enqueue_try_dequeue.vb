' vybe-test: vb/vb_concurrent_queue_operations/test_vb_concurrent_queue_enqueue_try_dequeue
' origin: languages/vb/tests/vb/test_vb_concurrent_queue_operations.rs

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
        Dim cq As New ConcurrentQueue(Of String)()
        cq.Enqueue("First")
        cq.Enqueue("Second")
        Dim item As String = Nothing
        Dim ok1 As Boolean = cq.TryDequeue(item)
        __Check(CStr(ok1), "True")
        __Check(CStr(item), "First")
        Dim ok2 As Boolean = cq.TryDequeue(item)
        __Check(CStr(ok2), "True")
        __Check(CStr(item), "Second")
    End Sub
End Module
