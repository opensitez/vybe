' vybe-test: vb/vb_async_cancellation_token/test_vb_async_cancellation_token_cancel_requested
' origin: languages/vb/tests/vb/test_vb_async_cancellation_token.rs

Imports System
Imports System.Threading
Imports System.Threading.Tasks

Module Program
    Async Function DoWorkAsync(token As CancellationToken) As Task
        If token.IsCancellationRequested Then
            Console.WriteLine("Canceled before start")
            Return
        End If
        Await Task.Delay(10, token)
    End Function

    Sub Main()
        Dim cts As New CancellationTokenSource()
        cts.Cancel()
        Try
            DoWorkAsync(cts.Token).Wait()
        Catch ex As Exception
            Console.WriteLine("Task Exception")
        End Try
    End Sub
End Module
