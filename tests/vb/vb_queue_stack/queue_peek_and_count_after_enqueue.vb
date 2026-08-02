' vybe-test: vb/vb_queue_stack/queue_peek_and_count_after_enqueue
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
        q.Enqueue(1)
        q.Enqueue(2)
        __Check(CStr(q.Peek()), "1")
        __Check(CStr(q.Count), "2")
    End Sub
End Module
