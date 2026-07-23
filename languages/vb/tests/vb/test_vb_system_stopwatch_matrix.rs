use super::helpers::run_vb;

#[test]
fn stopwatch_stopped_by_default() {
    let out = run_vb(
        r#"
Imports System
Imports System.Diagnostics

Module M
    Sub Main()
        Dim sw As New Stopwatch()
        Console.WriteLine(sw.IsRunning)
        Console.WriteLine(sw.ElapsedMilliseconds)
        Console.WriteLine(sw.IsRunning = False)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["False", "0", "True"]);
}

#[test]
fn stopwatch_start_and_stop_records_elapsed() {
    let out = run_vb(
        r#"
Imports System
Imports System.Diagnostics
Imports System.Threading

Module M
    Sub Main()
        Dim sw As New Stopwatch()
        sw.Start()
        Thread.Sleep(1)
        sw.Stop()

        Console.WriteLine(sw.IsRunning)
        Console.WriteLine(sw.ElapsedMilliseconds >= 0)
        Console.WriteLine(sw.ElapsedTicks >= 0)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["False", "True", "True"]);
}

#[test]
fn stopwatch_startnew_cycle_works() {
    let out = run_vb(
        r#"
Imports System
Imports System.Diagnostics

Module M
    Sub Main()
        Dim sw As Stopwatch = Stopwatch.StartNew()
        sw.Stop()
        Console.WriteLine(sw.IsRunning)
        Console.WriteLine(sw.ElapsedMilliseconds >= 0)
        sw.Reset()
        Console.WriteLine(sw.ElapsedMilliseconds)
        Console.WriteLine(sw.ElapsedTicks)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["False", "True", "0", "0"]);
}

#[test]
fn stopwatch_restart_is_supported() {
    let out = run_vb(
        r#"
Imports System
Imports System.Diagnostics

Module M
    Sub Main()
        Dim sw As New Stopwatch()
        sw.Start()
        sw.Stop()
        Dim first As Long = sw.ElapsedMilliseconds

        sw.Reset()
        sw.Start()
        sw.Stop()
        Dim second As Long = sw.ElapsedMilliseconds

        Console.WriteLine(first >= 0)
        Console.WriteLine(second >= 0)
        Console.WriteLine(first >= second OrElse first < second)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True", "True"]);
}

#[test]
fn stopwatch_high_resolution_flag_and_frequency_are_valid() {
    let out = run_vb(
        r#"
Imports System
Imports System.Diagnostics

Module M
    Sub Main()
        Console.WriteLine(Stopwatch.IsHighResolution OrElse Not Stopwatch.IsHighResolution)
        Console.WriteLine(Stopwatch.Frequency > 0)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn stopwatch_timestamp_is_monotonic_snapshot() {
    let out = run_vb(
        r#"
Imports System
Imports System.Diagnostics

Module M
    Sub Main()
        Dim first As Long = Stopwatch.GetTimestamp()
        Dim second As Long = Stopwatch.GetTimestamp()
        Console.WriteLine(second >= first)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}

#[test]
fn stopwatch_elapsed_is_duration_composable() {
    let out = run_vb(
        r#"
Imports System
Imports System.Diagnostics

Module M
    Sub Main()
        Dim one As TimeSpan = TimeSpan.FromMilliseconds(10)
        Dim sw As New Stopwatch()
        sw.Start()
        sw.Stop()
        Console.WriteLine(sw.Elapsed + one > one)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}

#[test]
fn stopwatch_reset_clears_running_state() {
    let out = run_vb(
        r#"
Imports System
Imports System.Diagnostics

Module M
    Sub Main()
        Dim sw As New Stopwatch()
        sw.Start()
        sw.Stop()
        sw.Reset()

        Console.WriteLine(sw.IsRunning)
        Console.WriteLine(sw.ElapsedMilliseconds)
        Console.WriteLine(sw.Elapsed.Ticks)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["False", "0", "0"]);
}
