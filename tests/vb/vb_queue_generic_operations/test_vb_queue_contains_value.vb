' vybe-test: vb/vb_queue_generic_operations/test_vb_queue_contains_value
' origin: languages/vb/tests/vb/test_vb_queue_generic_operations.rs

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

Module Program
    Sub Main()
        Dim q As New Queue(Of Integer)()
        q.Enqueue(100)
        q.Enqueue(200)
        __Check(CStr(q.Contains(100)), "True")
        __Check(CStr(q.Contains(300)), "False")
    End Sub
End Module
