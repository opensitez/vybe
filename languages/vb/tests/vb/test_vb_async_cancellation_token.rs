use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: CancellationTokenSource & Task Cancellation
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_async_cancellation_token_cancel_requested() {
    let src = r#"
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
"#;
    assert_eq!(run_vb(src), vec!["Canceled before start"]);
}

#[test]
fn test_vb_async_cancellation_token_source_cancel_after() {
    let src = r#"
Imports System.Threading

Module Program
    Sub Main()
        Dim cts As New CancellationTokenSource()
        cts.CancelAfter(50)
        Console.WriteLine(cts.IsCancellationRequested)
        Thread.Sleep(100)
        Console.WriteLine(cts.IsCancellationRequested)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False", "True"]);
}
