' vybe-test: vb/vb_concurrent_stack_push_pop/test_vb_concurrent_stack_multithreaded_push_pop
' origin: languages/vb/tests/vb/test_vb_concurrent_stack_push_pop.rs

Imports System.Collections.Concurrent
Imports System.Threading.Tasks

Module Program
    Sub Main()
        Dim s As New ConcurrentStack(Of Integer)()
        Parallel.For(0, 100, Sub(i) s.Push(i))
        Console.WriteLine("Stack Count: " & s.Count)
    End Sub
End Module
