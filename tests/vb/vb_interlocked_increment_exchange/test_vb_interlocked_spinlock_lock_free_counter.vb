' vybe-test: vb/vb_interlocked_increment_exchange/test_vb_interlocked_spinlock_lock_free_counter
' origin: languages/vb/tests/vb/test_vb_interlocked_increment_exchange.rs

Imports System.Threading
Imports System.Threading.Tasks

Module Program
    Sub Main()
        Dim counter As Integer = 0
        Dim tasks(9) As Task
        For i As Integer = 0 To 9
            tasks(i) = Task.Run(Sub()
                For j As Integer = 1 To 100
                    Interlocked.Increment(counter)
                Next
            End Sub)
        Next
        Task.WaitAll(tasks)
        Console.WriteLine(counter)
    End Sub
End Module
