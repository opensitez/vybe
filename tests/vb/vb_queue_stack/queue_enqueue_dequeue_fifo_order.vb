' vybe-test: vb/vb_queue_stack/queue_enqueue_dequeue_fifo_order
' origin: languages/vb/tests/vb/test_vb_queue_stack.rs

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

Imports System.Collections.Generic

Module M
    Sub Main()
        Dim q As New Queue(Of Integer)()
        q.Enqueue(10)
        q.Enqueue(20)
        __Check(CStr(q.Dequeue()), "10")
        __Check(CStr(q.Dequeue()), "20")
    End Sub
End Module
