' vybe-test: vb/vb_custom_event_thread_sync/test_vb_event_handler_multiple_parallel_triggers
' origin: languages/vb/tests/vb/test_vb_custom_event_thread_sync.rs

Imports System
Imports System.Threading.Tasks

Class ParallelEventSource
    Public Event Ping As EventHandler
    Public Sub Fire()
        RaiseEvent Ping(Me, EventArgs.Empty)
    End Sub
End Class

Module Program
    Sub Main()
        Dim pes As New ParallelEventSource()
        Dim counter = 0
        Dim lockObj As New Object()
        AddHandler pes.Ping, Sub(s, e)
            SyncLock lockObj
                counter += 1
            End SyncLock
        End Sub

        Parallel.For(0, 5, Sub(i) pes.Fire())
        Console.WriteLine("Parallel Count: " & counter)
    End Sub
End Module
