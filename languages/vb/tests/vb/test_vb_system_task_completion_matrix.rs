use super::helpers::run_vb;

#[test]
fn task_completion_result_is_observable() {
    let out = run_vb(
        r#"
Imports System
Imports System.Threading.Tasks

Module M
    Sub Main()
        Dim tcs As New TaskCompletionSource(Of Integer)()
        tcs.SetResult(42)

        Console.WriteLine(tcs.Task.IsCompleted)
        Console.WriteLine(tcs.Task.Result)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "42"]);
}

#[test]
fn task_completion_fail_path_reports_error() {
    let out = run_vb(
        r#"
Imports System
Imports System.Threading.Tasks

Module M
    Sub Main()
        Dim tcs As New TaskCompletionSource(Of Integer)()
        tcs.SetException(New InvalidOperationException("bad"))

        Try
            Console.WriteLine(tcs.Task.Result)
        Catch ex As AggregateException
            Console.WriteLine(ex.InnerException.GetType().Name)
        End Try
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["InvalidOperationException"]);
}

#[test]
fn task_completion_cancel_sets_status() {
    let out = run_vb(
        r#"
Imports System
Imports System.Threading
Imports System.Threading.Tasks

Module M
    Sub Main()
        Dim cts As New CancellationTokenSource()
        Dim task As Task = cts.Task

        cts.Cancel()

        Console.WriteLine(task.IsCanceled)
        Console.WriteLine(cts.IsCancellationRequested)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn task_completion_wait_for_all_from_factory_run() {
    let out = run_vb(
        r#"
Imports System
Imports System.Threading.Tasks

Module M
    Sub Main()
        Dim a As Task(Of Integer) = Task.Run(Function() 1)
        Dim b As Task(Of Integer) = Task.Run(Function() 2)
        Dim all As Task(Of Integer()) = Task.WhenAll(a, b)

        Console.WriteLine(all.IsCompleted)
        Console.WriteLine(all.Result.Sum())
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "3"]);
}

#[test]
fn task_completion_continuation_runs_on_completion() {
    let out = run_vb(
        r#"
Imports System
Imports System.Threading.Tasks

Module M
    Sub Main()
        Dim t As Task(Of Integer) = Task.Run(Function() 2)
        Dim continuation As Task(Of Integer) = t.ContinueWith(Function(x) x.Result + 1)

        Console.WriteLine(continuation.Result)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["3"]);
}
