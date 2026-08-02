' vybe-test: vb/vb_collections_advanced/queue_operations
' origin: languages/vb/tests/vb/test_vb_collections_advanced.rs

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

Module M
    Sub Main()
        Dim q As New Queue(Of String)
        q.Enqueue("first")
        q.Enqueue("second")
        q.Enqueue("third")
        __Check(CStr(q.Count), "3")
        __Check(CStr(q.Dequeue()), "first")
        __Check(CStr(q.Peek()), "second")
        __Check(CStr(q.Count), "2")
    End Sub
End Module
