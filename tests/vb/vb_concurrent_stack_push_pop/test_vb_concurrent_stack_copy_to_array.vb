' vybe-test: vb/vb_concurrent_stack_push_pop/test_vb_concurrent_stack_copy_to_array
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
        Dim s As New ConcurrentStack(Of Integer)()
        s.Push(10)
        s.Push(20)

        Dim dest(3) As Integer
        s.CopyTo(dest, 1)
        __Check(CStr(String.Join(",", dest)), "0,20,10,0")
    End Sub
End Module
