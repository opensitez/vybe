use super::helpers::run_vb;

#[test]
fn threadpool_get_available_threads_is_non_negative() {
    let out = run_vb(
        r#"
Imports System
Imports System.Threading

Module M
    Sub Main()
        Dim workerThreads As Integer
        Dim completionThreads As Integer
        ThreadPool.GetAvailableThreads(workerThreads, completionThreads)

        Console.WriteLine(workerThreads >= 0)
        Console.WriteLine(completionThreads >= 0)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn threadpool_get_min_threads_has_positive_values() {
    let out = run_vb(
        r#"
Imports System
Imports System.Threading

Module M
    Sub Main()
        Dim workerThreads As Integer
        Dim completionThreads As Integer
        ThreadPool.GetMinThreads(workerThreads, completionThreads)

        Console.WriteLine(workerThreads > 0)
        Console.WriteLine(completionThreads >= 0)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn threadpool_get_and_set_max_threads_round_trip() {
    let out = run_vb(
        r#"
Imports System
Imports System.Threading

Module M
    Sub Main()
        Dim workerThreads As Integer
        Dim completionThreads As Integer
        ThreadPool.GetMaxThreads(workerThreads, completionThreads)

        Dim reapply As Boolean = ThreadPool.SetMaxThreads(workerThreads, completionThreads)

        Console.WriteLine(reapply)
        Console.WriteLine(workerThreads > 0)
        Console.WriteLine(completionThreads >= 0)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True", "True"]);
}

#[test]
fn threadpool_requeue_of_work_item_runs() {
    let out = run_vb(
        r#"
Imports System
Imports System.Threading

Module M
    Sub Main()
        Dim done As New AutoResetEvent(False)
        Dim payload As String = "ok"

        ThreadPool.QueueUserWorkItem(
            Sub(_)
                payload = "done"
                done.Set()
            End Sub
        )

        Console.WriteLine(done.WaitOne(2000))
        Console.WriteLine(payload)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "done"]);
}

#[test]
fn threadpool_can_queue_multiple_items() {
    let out = run_vb(
        r#"
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
"#,
    );

    assert_eq!(out, vec!["True", "3"]);
}
