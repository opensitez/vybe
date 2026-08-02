' vybe-test: vb/vb_spec_arrays_collections/array_spec_queue_and_stack_can_model_fifo_and_lifo
' origin: languages/vb/tests/vb/test_vb_spec_arrays_collections.rs

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

Module M : Sub Main() : Dim q As New Queue(Of Integer) : q.Enqueue(1) : q.Enqueue(2) : Dim s As New Stack(Of Integer) : s.Push(1) : s.Push(2) : __Check(CStr(q.Dequeue()), "1") : __Check(CStr(s.Pop()), "2") : End Sub : End Module
