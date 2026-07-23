use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Task.WhenAll & Task.WhenAny Combinators
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_async_task_when_all() {
    let src = r#"
Imports System.Threading.Tasks

Module Program
    Async Function FetchOne() As Task(Of String)
        Await Task.Delay(5)
        Return "One"
    End Function

    Async Function FetchTwo() As Task(Of String)
        Await Task.Delay(5)
        Return "Two"
    End Function

    Async Function RunAllAsync() As Task
        Dim results As String() = Await Task.WhenAll(FetchOne(), FetchTwo())
        Console.WriteLine(String.Join(",", results))
    End Function

    Sub Main()
        RunAllAsync().Wait()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["One,Two"]);
}

#[test]
fn test_vb_async_task_when_any() {
    let src = r#"
Imports System.Threading.Tasks

Module Program
    Async Function SlowTask() As Task(Of String)
        Await Task.Delay(50)
        Return "Slow"
    End Function

    Async Function FastTask() As Task(Of String)
        Await Task.Delay(5)
        Return "Fast"
    End Function

    Async Function RunFirstAsync() As Task
        Dim winner As Task(Of String) = Await Task.WhenAny(SlowTask(), FastTask())
        Dim val As String = Await winner
        Console.WriteLine(val)
    End Function

    Sub Main()
        RunFirstAsync().Wait()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Fast"]);
}
