use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Async / Await & Task.Run
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_async_function_return_task_of_t() {
    let src = r#"
Imports System.Threading.Tasks

Module Program
    Async Function ComputeAsync() As Task(Of Integer)
        Await Task.Delay(10)
        Return 42
    End Function

    Sub Main()
        Dim t = ComputeAsync()
        Console.WriteLine(t.Result)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["42"]);
}

#[test]
fn test_vb_async_sub_void_lambda() {
    let src = r#"
Imports System.Threading.Tasks

Module Program
    Async Function ExecuteLambdaAsync() As Task
        Dim act As Func(Of Task) = Async Function()
            Await Task.Delay(10)
            Console.WriteLine("Lambda Executed")
        End Function
        Await act()
    End Function

    Sub Main()
        ExecuteLambdaAsync().Wait()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Lambda Executed"]);
}
