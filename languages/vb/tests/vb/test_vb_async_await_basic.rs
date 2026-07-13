use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Async and Await (Basic)
// ═══════════════════════════════════════════════════════════

#[test]
fn async_await_task_basic() {
    let out = run_vb(
        r#"
Imports System.Threading.Tasks

Module M
    Async Function DoWorkAsync() As Task
        Console.WriteLine("Start Work")
        Await Task.Delay(10)
        Console.WriteLine("End Work")
    End Function

    Sub Main()
        ' Wait synchronously for test validation
        DoWorkAsync().Wait()
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Start Work", "End Work"]);
}
