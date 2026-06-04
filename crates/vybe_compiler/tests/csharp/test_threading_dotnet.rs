use super::helpers::run_csharp;

#[test]
fn fully_qualified_thread_sleep_uses_shared_dotnet_surface() {
    let out = run_csharp(
        r#"
        Console.WriteLine("before");
        System.Threading.Thread.Sleep(1);
        Console.WriteLine("after");
    "#,
    );
    assert_eq!(out, vec!["before", "after"]);
}

#[test]
fn fully_qualified_process_start_info_wait_for_exit_uses_shared_dotnet_surface() {
    let out = run_csharp(
        r#"
        var si = new System.Diagnostics.ProcessStartInfo("/usr/bin/test", "hello = hello");
        var p = System.Diagnostics.Process.Start(si);
        p.WaitForExit();
        Console.WriteLine(p.ExitCode);
    "#,
    );
    assert_eq!(out, vec!["0"]);
}
