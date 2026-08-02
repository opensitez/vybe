' vybe-test: vb/vb_concurrent_stack_push_pop/test_vb_concurrent_stack_push_and_try_pop_lifo
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
        s.Push("First")
        s.Push("Second")

        Dim item1 As String = Nothing
        Dim item2 As String = Nothing
        Dim ok1 = s.TryPop(item1)
        Dim ok2 = s.TryPop(item2)

        __Check(CStr(ok1 & "|" & item1 & "|" & ok2 & "|" & item2), "True|Second|True|First")
    End Sub
End Module
