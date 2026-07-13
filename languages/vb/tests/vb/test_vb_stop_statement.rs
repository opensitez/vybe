use super::helpers::run_vb;

#[test]
fn stop_statement() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine("Before Stop")
        ' Stop acts as a breakpoint but execution can resume if a debugger is attached.
        ' Without a debugger, it might just break or do nothing depending on runtime.
        ' In some environments it throws an exception, but in .NET it usually calls System.Diagnostics.Debugger.Break()
        Stop
        Console.WriteLine("After Stop")
    End Sub
End Module
"#,
    );
    // Depending on the test environment, Stop might be ignored or might output
    // We'll just test that it parses correctly and outputs "Before Stop". "After Stop" may also be printed if ignored.
    assert_eq!(out, vec!["Before Stop", "After Stop"]);
}
