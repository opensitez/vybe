' vybe-test: vb/vb_async_task_delay_cancellation/test_vb_async_loop_cancellation_check
' origin: languages/vb/tests/vb/test_vb_async_task_delay_cancellation.rs

Imports System.Threading
Imports System.Threading.Tasks

Module Program
    Private Async Function LoopAsync(token As CancellationToken) As Task(Of Integer)
        Dim iterations As Integer = 0
        While Not token.IsCancellationRequested
            iterations += 1
            If iterations >= 3 Then Break
            Await Task.Delay(1, token)
        End While
        Return iterations
    End Function

    Sub Main()
        Dim cts As New CancellationTokenSource()
        Dim t = LoopAsync(cts.Token)
        Console.WriteLine(t.Result)
    End Sub
End Module
