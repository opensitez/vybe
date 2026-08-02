' vybe-test: vb/vb_async_configure_await_false/test_vb_async_configure_await_bypasses_sync_context
' origin: languages/vb/tests/vb/test_vb_async_configure_await_false.rs

Imports System.Threading
Imports System.Threading.Tasks

Class CustomSyncContext
    Inherits SynchronizationContext
    Public Overrides Sub Post(d As SendOrPostCallback, state As Object)
        Console.WriteLine("SyncContext Post Triggered")
        d(state)
    End Sub
End Class

Module Program
    Private Async Function BypassContextAsync() As Task
        ' ConfigureAwait(False) suppresses posting back to SyncContext!
        Await Task.Delay(5).ConfigureAwait(False)
    End Function

    Sub Main()
        Dim originalCtx = SynchronizationContext.Current
        Try
            SynchronizationContext.SetSynchronizationContext(New CustomSyncContext())
            Dim t = BypassContextAsync()
            t.Wait()
            Console.WriteLine("Bypass Async Finished")
        Finally
            SynchronizationContext.SetSynchronizationContext(originalCtx)
        End Try
    End Sub
End Module
