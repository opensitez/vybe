use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Async and Await (Task(Of T))
// ═══════════════════════════════════════════════════════════

#[test]
fn async_await_task_of_t() {
    let out = run_vb(
        r#"
Imports System.Threading.Tasks

Module M
    Async Function ComputeValueAsync() As Task(Of Integer)
        Await Task.Delay(10)
        Return 42
    End Function

    Sub Main()
        Dim t As Task(Of Integer) = ComputeValueAsync()
        Console.WriteLine(t.Result)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["42"]);
}
