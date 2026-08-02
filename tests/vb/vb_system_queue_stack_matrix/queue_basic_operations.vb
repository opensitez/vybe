' vybe-test: vb/vb_system_queue_stack_matrix/queue_basic_operations
' origin: languages/vb/tests/vb/test_vb_system_queue_stack_matrix.rs

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
        Dim queue As New Queue(Of Integer)()
        queue.Enqueue(1)
        queue.Enqueue(2)
        queue.Enqueue(3)
        __Check(CStr(queue.Count), "3")
        __Check(CStr(queue.Dequeue()), "1")
        __Check(CStr(queue.Peek()), "2")
        __Check(CStr(queue.Count), "2")
    End Sub
End Module
