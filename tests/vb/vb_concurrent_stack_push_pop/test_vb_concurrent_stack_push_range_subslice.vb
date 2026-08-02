' vybe-test: vb/vb_concurrent_stack_push_pop/test_vb_concurrent_stack_push_range_subslice
' origin: languages/vb/tests/vb/test_vb_concurrent_stack_push_pop.rs

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
        Dim s As New ConcurrentStack(Of String)()
        Dim raw As String() = {"A", "B", "C", "D"}
        ' PushRange(items, offset, count)
        s.PushRange(raw, 1, 2)

        Dim top As String = Nothing
        s.TryPop(top)
        __Check(CStr(top & "|Count=" & s.Count), "C|Count=1")
    End Sub
End Module
