' vybe-test: vb/vb_system_queue_stack_matrix/queue_contains_and_copyto
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
        Dim queue As New Queue(Of String)()
        queue.Enqueue("a")
        queue.Enqueue("b")
        queue.Enqueue("c")
        __Check(CStr(queue.Contains("b")), "True")
        Dim target(2) As String
        queue.CopyTo(target, 0)
        __Check(CStr(target(0)), "a")
        __Check(CStr(target(2)), "c")
    End Sub
End Module
