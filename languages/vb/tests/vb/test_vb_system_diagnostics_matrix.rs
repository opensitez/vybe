use super::helpers::run_vb;

#[test]
fn diagnostics_stopwatch_default_is_not_running() {
    let out = run_vb(
        r#"
Imports System.Diagnostics

Module M
    Sub Main()
        Dim sw As New Stopwatch()
        Console.WriteLine(sw.IsRunning)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["False"]);
}

#[test]
fn diagnostics_stopwatch_start_stop_cycle() {
    let out = run_vb(
        r#"
Imports System.Diagnostics

Module M
    Sub Main()
        Dim sw As Stopwatch = Stopwatch.StartNew()
        sw.Stop()
        Console.WriteLine(sw.IsRunning = False)
        Console.WriteLine(sw.ElapsedMilliseconds >= 0)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn diagnostics_stopwatch_timestamp_monotonic() {
    let out = run_vb(
        r#"
Imports System.Diagnostics

Module M
    Sub Main()
        Dim a As Long = Stopwatch.GetTimestamp()
        Dim b As Long = Stopwatch.GetTimestamp()
        Console.WriteLine(b >= a)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}

#[test]
fn diagnostics_stopwatch_reset_resets_state() {
    let out = run_vb(
        r#"
Imports System.Diagnostics

Module M
    Sub Main()
        Dim sw As New Stopwatch()
        sw.Start()
        sw.Stop()
        sw.Reset()
        Console.WriteLine(sw.IsRunning)
        Console.WriteLine(sw.ElapsedMilliseconds)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["False", "0"]);
}

#[test]
fn diagnostics_high_resolution_flag_is_boolean() {
    let out = run_vb(
        r#"
Imports System.Diagnostics

Module M
    Sub Main()
        Console.WriteLine(Stopwatch.IsHighResolution OrElse Not Stopwatch.IsHighResolution)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}

#[test]
fn diagnostics_frequency_is_positive() {
    let out = run_vb(
        r#"
Imports System.Diagnostics

Module M
    Sub Main()
        Console.WriteLine(Stopwatch.Frequency > 0)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}
