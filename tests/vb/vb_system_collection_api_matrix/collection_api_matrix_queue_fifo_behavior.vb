' vybe-test: vb/vb_system_collection_api_matrix/collection_api_matrix_queue_fifo_behavior
' origin: languages/vb/tests/vb/test_vb_system_collection_api_matrix.rs

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

        Dim a As String = queue.Dequeue()
        Dim b As String = queue.Dequeue()

        __Check(CStr(a), "a")
        __Check(CStr(b), "b")
        __Check(CStr(queue.Count), "1")
    End Sub
End Module
