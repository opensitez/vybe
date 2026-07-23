use super::helpers::run_vb;

#[test]
fn threading_timer_periodic_tick_happens() {
    let out = run_vb(
        r#"
Imports System
Imports System.Threading

Module M
    Private Shared _tickCount As Integer

    Sub Main()
        _tickCount = 0
        Dim done As New AutoResetEvent(False)

        Using timer As New Timer(New TimerCallback(AddressOf OnPeriodicTick), done, 0, 5)
            done.WaitOne(2000)
        End Using

        Console.WriteLine(_tickCount > 0)
        Console.WriteLine(_tickCount >= 1)
    End Sub

    Private Shared Sub OnPeriodicTick(state As Object)
        Interlocked.Increment(_tickCount)
        Dim gate As AutoResetEvent = CType(state, AutoResetEvent)
        gate.Set()
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn threading_timer_change_updates_schedule() {
    let out = run_vb(
        r#"
Imports System
Imports System.Threading

Module M
    Private Shared _tickCount As Integer

    Sub Main()
        _tickCount = 0
        Dim done As New AutoResetEvent(False)

        Dim timer As New Timer(New TimerCallback(AddressOf OnChangeTick), done, 80, Timeout.Infinite)
        Console.WriteLine(done.WaitOne(1) = False)

        timer.Change(0, Timeout.Infinite)
        Console.WriteLine(done.WaitOne(2000))

        timer.Dispose()
        Console.WriteLine(_tickCount > 0)
    End Sub

    Private Shared Sub OnChangeTick(state As Object)
        Interlocked.Increment(_tickCount)
        Dim gate As AutoResetEvent = CType(state, AutoResetEvent)
        gate.Set()
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True", "True"]);
}

#[test]
fn threading_timer_dispose_reclaims_one_shot_timer() {
    let out = run_vb(
        r#"
Imports System
Imports System.Threading

Module M
    Private Shared _tickCount As Integer

    Sub Main()
        _tickCount = 0
        Dim done As New AutoResetEvent(False)

        Dim timer As New Timer(New TimerCallback(AddressOf OnOneShotTick), done, 20, Timeout.Infinite)
        done.WaitOne(2000)
        timer.Dispose()

        Console.WriteLine(_tickCount > 0)
        Console.WriteLine(_tickCount = 1)
    End Sub

    Private Shared Sub OnOneShotTick(state As Object)
        Interlocked.Increment(_tickCount)
        Dim gate As AutoResetEvent = CType(state, AutoResetEvent)
        gate.Set()
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}
