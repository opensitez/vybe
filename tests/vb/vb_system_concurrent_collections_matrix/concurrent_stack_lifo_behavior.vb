' vybe-test: vb/vb_system_concurrent_collections_matrix/concurrent_stack_lifo_behavior
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
        Dim s As New ConcurrentStack(Of Integer)()
        s.Push(1)
        s.Push(2)

        Dim top As Integer = 0
        __Check(CStr(s.TryPeek(top)), "True")
        __Check(CStr(top), "2")

        Dim removed As Integer = 0
        __Check(CStr(s.TryPop(removed)), "True")
        __Check(CStr(removed), "2")
        __Check(CStr(s.Count), "1")
    End Sub
End Module
