' vybe-test: vb/vb_concurrent_stack_push_pop/test_vb_concurrent_stack_try_pop_range_array
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
        s.Push("A")
        s.Push("B")
        s.Push("C")

        Dim buffer(2) As String
        ' TryPopRange pops up to count elements into buffer array
        Dim poppedCount = s.TryPopRange(buffer, 0, 2)
        __Check(CStr(poppedCount & "|" & buffer(0) & "|" & buffer(1)), "2|C|B")
    End Sub
End Module
