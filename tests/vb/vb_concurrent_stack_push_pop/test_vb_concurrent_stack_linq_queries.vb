' vybe-test: vb/vb_concurrent_stack_push_pop/test_vb_concurrent_stack_linq_queries
' origin: languages/vb/tests/vb/test_vb_concurrent_stack_push_pop.rs

Imports System.Collections.Concurrent
Imports System.Linq

Module Program
    Sub Main()
        Dim s As New ConcurrentStack(Of Integer)()
        For i As Integer = 1 To 5 : s.Push(i) : Next
        Dim filtered = s.Where(Function(n) n > 2).ToList()
        Console.WriteLine(String.Join(",", filtered))
    End Sub
End Module
