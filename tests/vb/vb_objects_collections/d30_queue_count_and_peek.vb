' vybe-test: vb/vb_objects_collections/d30_queue_count_and_peek
' origin: languages/vb/tests/vb/vb_objects_collections_test.rs

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

Dim q As New Queue(Of String)
q.Enqueue("a")
q.Enqueue("b")
__Check(CStr(q.Count), "2")
__Check(CStr(q.Peek()), "a")
__Check(CStr(q.Count), "2")
