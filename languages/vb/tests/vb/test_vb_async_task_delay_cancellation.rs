use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Async Task.Delay & CancellationToken Mechanics
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_async_task_delay_basic_execution() {
    let src = r#"
Imports System.Threading.Tasks

Module Program
    Private Async Function RunDelayAsync() As Task(Of String)
        Await Task.Delay(10)
        Return "Delay Completed"
    End Function

    Sub Main()
        Dim t = RunDelayAsync()
        Console.WriteLine(t.Result)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Delay Completed"]);
}

#[test]
fn test_vb_async_cancellation_token_source_cancel() {
    let src = r#"
Imports System
Imports System.Threading
Imports System.Threading.Tasks

Module Program
    Private Async Function WorkAsync(token As CancellationToken) As Task(Of Boolean)
        Try
            Await Task.Delay(5000, token)
            Return True
        Catch ex As OperationCanceledException
            Console.WriteLine("Operation Canceled")
            Return False
        End Try
    End Function

    Sub Main()
        Dim cts As New CancellationTokenSource()
        Dim task = WorkAsync(cts.Token)
        cts.Cancel()
        Console.WriteLine("Task Result: " & task.Result)
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["Operation Canceled", "Task Result: False"]
    );
}

#[test]
fn test_vb_cancellation_token_throw_if_cancellation_requested() {
    let src = r#"
Imports System
Imports System.Threading

Module Program
    Private Sub CheckToken(token As CancellationToken)
        token.ThrowIfCancellationRequested()
    End Sub

    Sub Main()
        Dim cts As New CancellationTokenSource()
        cts.Cancel()
        Try
            CheckToken(cts.Token)
        Catch ex As OperationCanceledException
            Console.WriteLine("ThrowIfCancellationRequested Triggered")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["ThrowIfCancellationRequested Triggered"]);
}

#[test]
fn test_vb_cancellation_token_register_callback() {
    let src = r#"
Imports System.Threading

Module Program
    Sub Main()
        Dim cts As New CancellationTokenSource()
        cts.Token.Register(Sub() Console.WriteLine("Cancellation Callback Fired"))
        cts.Cancel()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Cancellation Callback Fired"]);
}

#[test]
fn test_vb_cancellation_token_source_cancel_after() {
    let src = r#"
Imports System
Imports System.Threading
Imports System.Threading.Tasks

Module Program
    Private Async Function DelayTask(token As CancellationToken) As Task
        Try
            Await Task.Delay(1000, token)
        Catch ex As OperationCanceledException
            Console.WriteLine("Timed Out Via CancelAfter")
        End Try
    End Function

    Sub Main()
        Dim cts As New CancellationTokenSource()
        cts.CancelAfter(10)
        Dim t = DelayTask(cts.Token)
        t.Wait()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Timed Out Via CancelAfter"]);
}

#[test]
fn test_vb_cancellation_token_none_is_uncancelable() {
    let src = r#"
Imports System.Threading

Module Program
    Sub Main()
        Dim token = CancellationToken.None
        Console.WriteLine(token.CanBeCanceled & "|" & token.IsCancellationRequested)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False|False"]);
}

#[test]
fn test_vb_cancellation_token_source_create_linked_token_source() {
    let src = r#"
Imports System
Imports System.Threading

Module Program
    Sub Main()
        Dim cts1 As New CancellationTokenSource()
        Dim cts2 As New CancellationTokenSource()
        Dim linked = CancellationTokenSource.CreateLinkedTokenSource(cts1.Token, cts2.Token)

        AddHandler linked.Token.Register, Sub() Console.WriteLine("Linked Token Canceled")
        cts2.Cancel()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Linked Token Canceled"]);
}

#[test]
fn test_vb_async_task_delay_timespan_overload() {
    let src = r#"
Imports System
Imports System.Threading.Tasks

Module Program
    Private Async Function RunTimeSpanDelayAsync() As Task(Of String)
        Await Task.Delay(TimeSpan.FromMilliseconds(10))
        Return "TimeSpan Delay Passed"
    End Function

    Sub Main()
        Dim t = RunTimeSpanDelayAsync()
        Console.WriteLine(t.Result)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["TimeSpan Delay Passed"]);
}

#[test]
fn test_vb_async_task_when_all_multiple_delays() {
    let src = r#"
Imports System.Threading.Tasks

Module Program
    Private Async Function RunParallelDelaysAsync() As Task(Of String)
        Dim t1 = Task.Delay(10)
        Dim t2 = Task.Delay(20)
        Await Task.WhenAll(t1, t2)
        Return "All Delays Finished"
    End Function

    Sub Main()
        Dim t = RunParallelDelaysAsync()
        Console.WriteLine(t.Result)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["All Delays Finished"]);
}

#[test]
fn test_vb_async_task_when_any_first_delay_wins() {
    let src = r#"
Imports System.Threading.Tasks

Module Program
    Private Async Function RunRaceAsync() As Task(Of String)
        Dim tFast = Task.Run(Function() As String
            Task.Delay(5).Wait()
            Return "Fast"
        End Function)
        Dim tSlow = Task.Run(Function() As String
            Task.Delay(500).Wait()
            Return "Slow"
        End Function)

        Dim winner = Await Task.WhenAny(tFast, tSlow)
        Return Await winner
    End Function

    Sub Main()
        Dim t = RunRaceAsync()
        Console.WriteLine(t.Result)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Fast"]);
}

#[test]
fn test_vb_async_function_returning_value_task() {
    let src = r#"
Imports System.Threading.Tasks

Module Program
    Private Async Function GetCachedValueAsync(id As Integer) As ValueTask(Of String)
        If id = 1 Then Return "Cached"
        Await Task.Delay(5)
        Return "Computed"
    End Function

    Sub Main()
        Dim t1 = GetCachedValueAsync(1).AsTask()
        Dim t2 = GetCachedValueAsync(2).AsTask()
        Console.WriteLine(t1.Result & "|" & t2.Result)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Cached|Computed"]);
}

#[test]
fn test_vb_cancellation_token_is_cancellation_requested_property() {
    let src = r#"
Imports System.Threading

Module Program
    Sub Main()
        Dim cts As New CancellationTokenSource()
        Console.WriteLine(cts.Token.IsCancellationRequested)
        cts.Cancel()
        Console.WriteLine(cts.Token.IsCancellationRequested)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False", "True"]);
}

#[test]
fn test_vb_async_void_sub_exception_captured() {
    let src = r#"
Imports System.Threading.Tasks

Module Program
    Private Async Function SafeExecuteAsync() As Task
        Await Task.Yield()
        Console.WriteLine("Yield Completed")
    End Function

    Sub Main()
        Dim t = SafeExecuteAsync()
        t.Wait()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Yield Completed"]);
}

#[test]
fn test_vb_async_task_yield_reschedules() {
    let src = r#"
Imports System.Threading.Tasks

Module Program
    Private Async Function YieldStepAsync() As Task(Of String)
        Console.WriteLine("Before Yield")
        Await Task.Yield()
        Console.WriteLine("After Yield")
        Return "Done"
    End Function

    Sub Main()
        Dim t = YieldStepAsync()
        Console.WriteLine(t.Result)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Before Yield", "After Yield", "Done"]);
}

#[test]
fn test_vb_cancellation_token_wait_handle() {
    let src = r#"
Imports System.Threading

Module Program
    Sub Main()
        Dim cts As New CancellationTokenSource()
        Dim handle = cts.Token.WaitHandle
        Console.WriteLine(handle IsNot Nothing)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_async_loop_cancellation_check() {
    let src = r#"
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
"#;
    assert_eq!(run_vb(src), vec!["3"]);
}

#[test]
fn test_vb_cancellation_token_source_dispose() {
    let src = r#"
Imports System.Threading

Module Program
    Sub Main()
        Using cts As New CancellationTokenSource()
            cts.Cancel()
            Console.WriteLine(cts.IsCancellationRequested)
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_async_task_delay_negative_one_infinite_delay() {
    let src = r#"
Imports System.Threading
Imports System.Threading.Tasks

Module Program
    Sub Main()
        Dim cts As New CancellationTokenSource()
        Dim t = Task.Delay(Timeout.Infinite, cts.Token)
        cts.Cancel()
        Console.WriteLine(t.IsCanceled)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_async_try_catch_around_delay() {
    let src = r#"
Imports System
Imports System.Threading
Imports System.Threading.Tasks

Module Program
    Private Async Function DelayedActionAsync(cts As CancellationTokenSource) As Task
        Try
            cts.Cancel()
            Await Task.Delay(1000, cts.Token)
        Catch ex As TaskCanceledException
            Console.WriteLine("TaskCanceledException Handled")
        End Try
    End Function

    Sub Main()
        Dim cts As New CancellationTokenSource()
        Dim t = DelayedActionAsync(cts)
        t.Wait()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["TaskCanceledException Handled"]);
}

#[test]
fn test_vb_cancellation_token_multiple_registrations() {
    let src = r#"
Imports System.Threading

Module Program
    Sub Main()
        Dim cts As New CancellationTokenSource()
        cts.Token.Register(Sub() Console.WriteLine("Reg 1"))
        cts.Token.Register(Sub() Console.WriteLine("Reg 2"))
        cts.Cancel()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Reg 2", "Reg 1"]);
}
