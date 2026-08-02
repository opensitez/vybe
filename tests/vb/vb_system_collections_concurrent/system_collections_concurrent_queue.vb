' vybe-test: vb/vb_system_collections_concurrent/system_collections_concurrent_queue
' origin: languages/vb/tests/vb/test_vb_system_collections_concurrent.rs

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
        Dim cq As New ConcurrentQueue(Of Integer)()
        
        cq.Enqueue(10)
        cq.Enqueue(20)
        
        Dim result As Integer
        If cq.TryDequeue(result) Then
            __Check(CStr(result), "10")
        End If
        
        __Check(CStr(cq.Count), "1")
    End Sub
End Module
