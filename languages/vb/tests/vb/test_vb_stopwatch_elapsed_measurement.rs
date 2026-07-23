use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: System.Diagnostics.Stopwatch Measurement
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_stopwatch_start_stop() {
    let src = r#"
Imports System.Diagnostics
Imports System.Threading

Module Program
    Sub Main()
        Dim sw As Stopwatch = Stopwatch.StartNew()
        Thread.Sleep(10)
        sw.Stop()
        Console.WriteLine(sw.IsRunning)
        Console.WriteLine(sw.ElapsedMilliseconds >= 5)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False", "True"]);
}
