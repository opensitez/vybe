use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Await in Catch/Finally blocks
// ═══════════════════════════════════════════════════════════

#[test]
fn async_await_in_catch_finally() {
    let out = run_vb(
        r#"
Imports System.Threading.Tasks

Module M
    Async Function LogErrorAsync() As Task
        Await Task.Delay(10)
        Console.WriteLine("Error Logged Async")
    End Function

    Async Function CleanupAsync() As Task
        Await Task.Delay(10)
        Console.WriteLine("Cleaned Up Async")
    End Function

    Async Function DoWorkAsync() As Task
        Try
            Throw New Exception("Fail")
        Catch ex As Exception
            Await LogErrorAsync()
        Finally
            Await CleanupAsync()
        End Try
    End Function

    Sub Main()
        DoWorkAsync().Wait()
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Error Logged Async", "Cleaned Up Async"]);
}
