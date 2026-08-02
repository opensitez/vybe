' vybe-test: vb/vb_queue_stack/queue_copy_from_array_roundtrip
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
        Dim source As Integer() = {9, 8, 7}
        Dim q As New Queue(Of Integer)(source)
        __Check(CStr(q.Count), "3")
        __Check(CStr(q.Dequeue()), "9")
        __Check(CStr(q.Dequeue()), "8")
        __Check(CStr(q.Dequeue()), "7")
    End Sub
End Module
