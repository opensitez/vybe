use super::helpers::run_vb;

#[test]
fn lock_monitor_matrix_synclock_prevents_race() {
    let out = run_vb(
        r#"
Imports System
Imports System.Threading

Module M
    Sub Main()
        Dim lockObject As New Object()
        Dim value As Integer = 0

        Dim t1 As New Thread(
            Sub()
                For i As Integer = 0 To 999
                    SyncLock lockObject
                        value += 1
                    End SyncLock
                Next
            End Sub)
        Dim t2 As New Thread(
            Sub()
                For i As Integer = 0 To 999
                    SyncLock lockObject
                        value += 1
                    End SyncLock
                Next
            End Sub)

        t1.Start()
        t2.Start()
        t1.Join()
        t2.Join()

        Console.WriteLine(value)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["2000"]);
}

#[test]
fn lock_monitor_matrix_sync_lock_reentrant() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim lockObject As New Object()
        Dim value As Integer = 0

        SyncLock lockObject
            value += 1
            SyncLock lockObject
                value += 2
            End SyncLock
            value += 1
        End SyncLock

        Console.WriteLine(value)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["4"]);
}

#[test]
fn lock_monitor_matrix_monitor_try_enter_contract() {
    let out = run_vb(
        r#"
Imports System.Threading

Module M
    Sub Main()
        Dim lockObject As New Object()
        Dim entered As Boolean = False

        entered = Monitor.TryEnter(lockObject)
        If entered Then
            Monitor.Exit(lockObject)
        End If

        Console.WriteLine(entered)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}

#[test]
fn lock_monitor_matrix_nested_lock_scope_counts() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim lockObject As New Object()
        Dim value As Integer = 0

        SyncLock lockObject
            value = 1
            SyncLock lockObject
                value = value + 1
            End SyncLock
            value = value * 2
        End SyncLock

        Console.WriteLine(value)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["4"]);
}
