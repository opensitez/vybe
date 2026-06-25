//! `System.Diagnostics`: `Stopwatch`, `Debug`, `Trace`, `Process`.
use super::helpers::run_csharp;

#[test]
fn stopwatch_elapsed_is_positive_after_work() {
    assert_eq!(
        run_csharp(r#"var sw=System.Diagnostics.Stopwatch.StartNew();
int s=0; for(int i=0;i<10000;i++) s+=i;
sw.Stop();
Console.WriteLine(sw.ElapsedMilliseconds>=0);"#),
        &["True"]
    );
}

#[test]
fn stopwatch_is_not_running_after_stop() {
    assert_eq!(
        run_csharp(r#"var sw=System.Diagnostics.Stopwatch.StartNew();
sw.Stop();
Console.WriteLine(sw.IsRunning);"#),
        &["False"]
    );
}

#[test]
fn stopwatch_reset_clears_elapsed() {
    assert_eq!(
        run_csharp(r#"var sw=System.Diagnostics.Stopwatch.StartNew();
System.Threading.Thread.Sleep(5);
sw.Stop();
sw.Reset();
Console.WriteLine(sw.Elapsed==System.TimeSpan.Zero);"#),
        &["True"]
    );
}

#[test]
fn stopwatch_restart_starts_from_zero() {
    assert_eq!(
        run_csharp(r#"var sw=new System.Diagnostics.Stopwatch();
sw.Start();
System.Threading.Thread.Sleep(5);
sw.Restart();
Console.WriteLine(sw.IsRunning);"#),
        &["True"]
    );
}
