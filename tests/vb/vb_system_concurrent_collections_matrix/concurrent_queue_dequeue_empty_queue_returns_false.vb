' vybe-test: vb/vb_system_concurrent_collections_matrix/concurrent_queue_dequeue_empty_queue_returns_false
' origin: languages/vb/tests/vb/test_vb_system_concurrent_collections_matrix.rs

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

Module M
    Sub Main()
        Dim q As New ConcurrentQueue(Of Integer)()
        Dim value As Integer = 42

        __Check(CStr(q.TryDequeue(value)), "False")
        __Check(CStr(value), "42")
        __Check(CStr(q.Count), "0")
    End Sub
End Module
