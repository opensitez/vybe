' vybe-test: vb/vb_concurrent_stack_operations/test_vb_concurrent_stack_push_range_try_pop_range
' origin: languages/vb/tests/vb/test_vb_concurrent_stack_operations.rs

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

Module Program
    Sub Main()
        Dim cs As New ConcurrentStack(Of Integer)()
        Dim items As Integer() = {1, 2, 3, 4}
        cs.PushRange(items)
        Dim popped(1) As Integer
        Dim count As Integer = cs.TryPopRange(popped)
        __Check(CStr(count), "2")
        __Check(CStr(String.Join(",", popped)), "4,3")
    End Sub
End Module
