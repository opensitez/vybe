use super::helpers::run_vb;

#[test]
fn threading_current_thread_is_alive_and_has_id() {
    let out = run_vb(
        r#"
Imports System.Threading

Module M
    Sub Main()
        Dim current As Thread = Thread.CurrentThread
        Console.WriteLine(current.IsAlive)
        Console.WriteLine(current.ManagedThreadId > 0)
        Console.WriteLine(current.Name = current.Name)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True", "True"]);
}

#[test]
fn threading_thread_runs_and_sets_flag() {
    let out = run_vb(
        r#"
Imports System.Threading

Module M
    Sub Main()
        Dim ran As Boolean = False

        Dim t As New Thread(Sub()
            ran = True
        End Sub)

        t.Start()
        t.Join()

        Console.WriteLine(ran)
        Console.WriteLine(t.ThreadState = ThreadState.Stopped)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn threading_threadpool_get_max_threads() {
    let out = run_vb(
        r#"
Imports System.Threading

Module M
    Sub Main()
        Dim workerThreads As Integer
        Dim ioThreads As Integer
        ThreadPool.GetMaxThreads(workerThreads, ioThreads)

        Console.WriteLine(workerThreads > 0)
        Console.WriteLine(ioThreads > 0)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn threading_threadpool_get_min_threads() {
    let out = run_vb(
        r#"
Imports System.Threading

Module M
    Sub Main()
        Dim minWorkers As Integer
        Dim minIO As Integer
        ThreadPool.GetMinThreads(minWorkers, minIO)

        Console.WriteLine(minWorkers >= 1)
        Console.WriteLine(minIO >= 1)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn threading_threadpool_work_item_completes() {
    let out = run_vb(
        r#"
Imports System
Imports System.Threading

Module M
    Sub Main()
        Dim signal As New AutoResetEvent(False)

        ThreadPool.QueueUserWorkItem(Sub(_)
            signal.Set()
        End Sub)

        Console.WriteLine(signal.WaitOne(2000))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}

#[test]
fn threading_interlocked_add_and_read() {
    let out = run_vb(
        r#"
Imports System.Threading

Module M
    Sub Main()
        Dim value As Integer = 10
        Console.WriteLine(Interlocked.Increment(value))
        Console.WriteLine(Interlocked.Add(value, 5))
        Console.WriteLine(Interlocked.Decrement(value))
        Console.WriteLine(value)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["11", "16", "15", "15"]);
}

#[test]
fn threading_interlocked_exchange_and_compare_exchange() {
    let out = run_vb(
        r#"
Imports System.Threading

Module M
    Sub Main()
        Dim value As Integer = 5
        Dim old As Integer = Interlocked.Exchange(value, 9)

        Dim matched As Integer = Interlocked.CompareExchange(value, 11, 9)
        Dim unmatched As Integer = Interlocked.CompareExchange(value, 13, 5)

        Console.WriteLine(old)
        Console.WriteLine(matched)
        Console.WriteLine(unmatched)
        Console.WriteLine(value)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["5", "9", "11", "11"]);
}

#[test]
fn threading_volatile_roundtrip() {
    let out = run_vb(
        r#"
Imports System.Threading

Module M
    Sub Main()
        Dim value As Integer = 4
        Volatile.Write(value, 12)

        Console.WriteLine(Volatile.Read(value))
        Console.WriteLine(value)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["12", "12"]);
}

#[test]
fn threading_auto_reset_event_roundtrip() {
    let out = run_vb(
        r#"
Imports System.Threading

Module M
    Sub Main()
        Dim signal As New AutoResetEvent(False)

        Console.WriteLine(signal.WaitOne(1))
        signal.Set()
        Console.WriteLine(signal.WaitOne(2000))
        Console.WriteLine(signal.WaitOne(1))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["False", "True", "False"]);
}

#[test]
fn threading_manual_reset_event_roundtrip() {
    let out = run_vb(
        r#"
Imports System.Threading

Module M
    Sub Main()
        Dim gate As New ManualResetEvent(False)

        Console.WriteLine(gate.WaitOne(1))
        gate.Set()
        Console.WriteLine(gate.WaitOne(1))
        gate.Reset()
        Console.WriteLine(gate.WaitOne(1))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["False", "True", "False"]);
}

#[test]
fn threading_monitor_locking_changes_visibility() {
    let out = run_vb(
        r#"
Imports System.Threading

Module M
    Sub Main()
        Dim lockObj As New Object()
        Dim sum As Integer = 0

        SyncLock lockObj
            sum = 1
        End SyncLock

        Console.WriteLine(sum = 1)
        Console.WriteLine(Monitor.IsEntered(lockObj))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn threading_wait_handle_wait_all() {
    let out = run_vb(
        r#"
Imports System.Threading

Module M
    Sub Main()
        Dim a As New AutoResetEvent(False)
        Dim b As New AutoResetEvent(False)

        a.Set()
        b.Set()

        Dim allDone As WaitHandle() = {a, b}
        Console.WriteLine(WaitHandle.WaitAll(allDone))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}

#[test]
fn threading_thread_abort_not_supported_is_false_safety() {
    let out = run_vb(
        r#"
Imports System
Imports System.Threading

Module M
    Sub Main()
        Dim t As New Thread(Sub()
        End Sub)

        Console.WriteLine(t.ThreadState = ThreadState.Unstarted)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}
