use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Async ConfigureAwait & SynchronizationContext
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_async_configure_await_false_execution() {
    let src = r#"
Imports System.Threading.Tasks

Module Program
    Private Async Function RunConfigureAwaitFalseAsync() As Task(Of String)
        Await Task.Delay(10).ConfigureAwait(False)
        Return "ConfigureAwait(False) Completed"
    End Function

    Sub Main()
        Dim t = RunConfigureAwaitFalseAsync()
        Console.WriteLine(t.Result)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["ConfigureAwait(False) Completed"]);
}

#[test]
fn test_vb_async_configure_await_true_execution() {
    let src = r#"
Imports System.Threading.Tasks

Module Program
    Private Async Function RunConfigureAwaitTrueAsync() As Task(Of String)
        Await Task.Delay(10).ConfigureAwait(True)
        Return "ConfigureAwait(True) Completed"
    End Function

    Sub Main()
        Dim t = RunConfigureAwaitTrueAsync()
        Console.WriteLine(t.Result)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["ConfigureAwait(True) Completed"]);
}

#[test]
fn test_vb_async_configure_await_chained_calls() {
    let src = r#"
Imports System.Threading.Tasks

Module Program
    Private Async Function ChainedAsync() As Task(Of Integer)
        Dim val = 10
        Await Task.Delay(5).ConfigureAwait(False)
        val += 20
        Await Task.Delay(5).ConfigureAwait(False)
        val += 30
        Return val
    End Function

    Sub Main()
        Dim t = ChainedAsync()
        Console.WriteLine(t.Result)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["60"]);
}

#[test]
fn test_vb_async_configure_await_exception_handling() {
    let src = r#"
Imports System
Imports System.Threading.Tasks

Module Program
    Private Async Function FaultyAsync() As Task
        Await Task.Delay(5).ConfigureAwait(False)
        Throw New InvalidOperationException("Failure after await")
    End Function

    Sub Main()
        Try
            Dim t = FaultyAsync()
            t.Wait()
        Catch ex As AggregateException
            Console.WriteLine(ex.InnerException.Message)
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Failure after await"]);
}

#[test]
fn test_vb_async_task_run_configure_await() {
    let src = r#"
Imports System.Threading.Tasks

Module Program
    Private Async Function CalculateAsync() As Task(Of Integer)
        Dim res = Await Task.Run(Function() 42).ConfigureAwait(False)
        Return res * 2
    End Function

    Sub Main()
        Dim t = CalculateAsync()
        Console.WriteLine(t.Result)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["84"]);
}

#[test]
fn test_vb_async_valuetask_configure_await() {
    let src = r#"
Imports System.Threading.Tasks

Module Program
    Private Async Function GetValAsync() As ValueTask(Of String)
        Await Task.Yield().ConfigureAwait(False)
        Return "ValueTask Success"
    End Function

    Sub Main()
        Dim vt = GetValAsync()
        Console.WriteLine(vt.AsTask().Result)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["ValueTask Success"]);
}

#[test]
fn test_vb_async_synchronization_context_preservation() {
    let src = r#"
Imports System.Threading
Imports System.Threading.Tasks

Class DummySyncContext
    Inherits SynchronizationContext
    Public Overrides Sub Post(d As SendOrPostCallback, state As Object)
        Console.WriteLine("SyncContext Post Called")
        d(state)
    End Sub
End Class

Module Program
    Private Async Function ContextAsync() As Task
        Await Task.Yield()
    End Function

    Sub Main()
        Dim originalCtx = SynchronizationContext.Current
        Try
            SynchronizationContext.SetSynchronizationContext(New DummySyncContext())
            Dim t = ContextAsync()
            t.Wait()
        Finally
            SynchronizationContext.SetSynchronizationContext(originalCtx)
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["SyncContext Post Called"]);
}

#[test]
fn test_vb_async_configure_await_bypasses_sync_context() {
    let src = r#"
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
"#;
    assert_eq!(run_vb(src), vec!["Bypass Async Finished"]);
}

#[test]
fn test_vb_async_yield_configure_await() {
    let src = r#"
Imports System.Threading.Tasks

Module Program
    Private Async Function YieldConfiguredAsync() As Task(Of String)
        Await Task.Yield().ConfigureAwait(False)
        Return "Yield Configured Completed"
    End Function

    Sub Main()
        Dim t = YieldConfiguredAsync()
        Console.WriteLine(t.Result)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Yield Configured Completed"]);
}

#[test]
fn test_vb_async_multiple_await_mixed_configure_await() {
    let src = r#"
Imports System.Threading.Tasks

Module Program
    Private Async Function MixedAwaitsAsync() As Task(Of String)
        Await Task.Delay(5).ConfigureAwait(False)
        Await Task.Delay(5).ConfigureAwait(True)
        Return "Mixed Awaits Passed"
    End Function

    Sub Main()
        Dim t = MixedAwaitsAsync()
        Console.WriteLine(t.Result)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Mixed Awaits Passed"]);
}

#[test]
fn test_vb_async_configure_await_inside_loop() {
    let src = r#"
Imports System.Threading.Tasks

Module Program
    Private Async Function LoopAsync() As Task(Of Integer)
        Dim total = 0
        For i As Integer = 1 To 3
            Await Task.Delay(2).ConfigureAwait(False)
            total += i
        Next
        Return total
    End Function

    Sub Main()
        Dim t = LoopAsync()
        Console.WriteLine(t.Result)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["6"]);
}

#[test]
fn test_vb_async_configure_await_in_try_finally() {
    let src = r#"
Imports System.Threading.Tasks

Module Program
    Private Async Function ResourceAsync() As Task(Of String)
        Dim res = "Opened"
        Try
            Await Task.Delay(5).ConfigureAwait(False)
            Return res & " -> Processed"
        Finally
            Console.WriteLine("Finally Block Executed")
        End Try
    End Function

    Sub Main()
        Dim t = ResourceAsync()
        Console.WriteLine(t.Result)
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["Finally Block Executed", "Opened -> Processed"]
    );
}

#[test]
fn test_vb_async_configure_await_with_cancellation_token() {
    let src = r#"
Imports System.Threading
Imports System.Threading.Tasks

Module Program
    Private Async Function CancellableConfiguredAsync(cts As CancellationTokenSource) As Task(Of String)
        Await Task.Delay(10, cts.Token).ConfigureAwait(False)
        Return "Success"
    End Function

    Sub Main()
        Dim cts As New CancellationTokenSource()
        Dim t = CancellableConfiguredAsync(cts)
        Console.WriteLine(t.Result)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Success"]);
}

#[test]
fn test_vb_async_configure_await_nested_methods() {
    let src = r#"
Imports System.Threading.Tasks

Module Program
    Private Async Function InnerAsync() As Task(Of String)
        Await Task.Delay(5).ConfigureAwait(False)
        Return "Inner"
    End Function

    Private Async Function OuterAsync() As Task(Of String)
        Dim res = Await InnerAsync().ConfigureAwait(False)
        Return "Outer -> " & res
    End Function

    Sub Main()
        Dim t = OuterAsync()
        Console.WriteLine(t.Result)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Outer -> Inner"]);
}

#[test]
fn test_vb_async_configure_await_returning_tuple() {
    let src = r#"
Imports System.Threading.Tasks

Module Program
    Private Async Function GetTupleAsync() As Task(Of (Code As Integer, Message As String))
        Await Task.Delay(5).ConfigureAwait(False)
        Return (200, "SuccessTuple")
    End Function

    Sub Main()
        Dim t = GetTupleAsync()
        Console.WriteLine(t.Result.Code & " " & t.Result.Message)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["200 SuccessTuple"]);
}

#[test]
fn test_vb_async_configure_await_returning_anonymous_type() {
    let src = r#"
Imports System.Threading.Tasks

Module Program
    Private Async Function GetAnonAsync() As Task(Of Object)
        Await Task.Delay(5).ConfigureAwait(False)
        Return New With {.Tag = "ConfiguredAnon"}
    End Function

    Sub Main()
        Dim t = GetAnonAsync()
        Dim res As Dynamic = t.Result
        Console.WriteLine(res.Tag)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["ConfiguredAnon"]);
}

#[test]
fn test_vb_async_configure_await_generic_class() {
    let src = r#"
Imports System.Threading.Tasks

Class AsyncWorker(Of T)
    Public Async Function ProcessAsync(input As T) As Task(Of T)
        Await Task.Delay(5).ConfigureAwait(False)
        Return input
    End Function
End Class

Module Program
    Sub Main()
        Dim w As New AsyncWorker(Of String)()
        Dim t = w.ProcessAsync("GenericInput")
        Console.WriteLine(t.Result)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["GenericInput"]);
}

#[test]
fn test_vb_async_configure_await_event_handler_simulation() {
    let src = r#"
Imports System.Threading.Tasks

Module Program
    Private Async Function OnCustomEventAsync() As Task
        Await Task.Delay(5).ConfigureAwait(False)
        Console.WriteLine("Async Event Handler Completed")
    End Function

    Sub Main()
        Dim t = OnCustomEventAsync()
        t.Wait()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Async Event Handler Completed"]);
}

#[test]
fn test_vb_async_configure_await_with_already_completed_task() {
    let src = r#"
Imports System.Threading.Tasks

Module Program
    Private Async Function FastAsync() As Task(Of String)
        Dim completedTask = Task.FromResult("Instant")
        Dim res = Await completedTask.ConfigureAwait(False)
        Return res
    End Function

    Sub Main()
        Dim t = FastAsync()
        Console.WriteLine(t.Result)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Instant"]);
}

#[test]
fn test_vb_async_configure_await_struct_return() {
    let src = r#"
Imports System.Threading.Tasks

Structure Point
    Public X As Integer
    Public Y As Integer
    Public Sub New(x As Integer, y As Integer) : Me.X = x : Me.Y = y : End Sub
End Structure

Module Program
    Private Async Function GetPointAsync() As Task(Of Point)
        Await Task.Delay(5).ConfigureAwait(False)
        Return New Point(100, 200)
    End Function

    Sub Main()
        Dim t = GetPointAsync()
        Console.WriteLine(t.Result.X & "," & t.Result.Y)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["100,200"]);
}
