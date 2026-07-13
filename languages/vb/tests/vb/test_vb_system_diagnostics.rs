use super::helpers::run_vb;

#[test]
fn system_diagnostics_stopwatch() {
    let out = run_vb(
        r#"
Imports System.Diagnostics

Module M
    Sub Main()
        Dim sw As New Stopwatch()
        sw.Start()
        ' Simulate some work
        Dim sum = 0
        For i = 1 To 1000
            sum += i
        Next
        sw.Stop()
        
        Console.WriteLine(sw.IsRunning)
        Console.WriteLine(sw.ElapsedMilliseconds >= 0)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["False", "True"]);
}

#[test]
fn system_diagnostics_process() {
    let out = run_vb(
        r#"
Imports System.Diagnostics

Module M
    Sub Main()
        Dim p As Process = Process.GetCurrentProcess()
        Console.WriteLine(p IsNot Nothing)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["True"]);
}
