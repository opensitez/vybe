' vybe-test: vb/vb_concurrent_queue_enqueue_dequeue/test_vb_concurrent_queue_struct_elements
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

Structure TaskRecord
    Public Id As Integer
    Public TaskName As String
End Structure

Module Program
    Sub Main()
        Dim q As New ConcurrentQueue(Of TaskRecord)()
        q.Enqueue(New TaskRecord With {.Id = 1, .TaskName = "Job1"})

        Dim rec As TaskRecord
        q.TryDequeue(rec)
        __Check(CStr(rec.Id & ":" & rec.TaskName), "1:Job1")
    End Sub
End Module
