' vybe-test: vb/vb_system_threading_pool_matrix/threadpool_can_queue_multiple_items
' origin: languages/vb/tests/vb/test_vb_system_threading_pool_matrix.rs

Imports System
Imports System.Threading

Module M
    Sub Main()
        Dim done As Integer = 0
        Dim barrier As New CountdownEvent(3)

        Dim mark As New AutoResetEvent(False)

        Dim submit As Integer = 0
        For i As Integer = 1 To 3
            ThreadPool.QueueUserWorkItem(
                Sub(_)
                    Interlocked.Increment(done)
                    barrier.Signal()
                    If barrier.CurrentCount = 0 Then
                        mark.Set()
                    End If
                End Sub
            )
            submit += 1
        Next

        mark.WaitOne(2000)
        Console.WriteLine(done = submit)
        Console.WriteLine(done)
    End Sub
End Module
