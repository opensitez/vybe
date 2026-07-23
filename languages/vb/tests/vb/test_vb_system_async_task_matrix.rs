use super::helpers::run_vb;

#[test]
fn async_task_matrix_from_result_and_get_result() {
    let out = run_vb(
        r#"
Imports System
Imports System.Threading.Tasks

Module M
    Sub Main()
        Dim t As Task(Of Integer) = Task.FromResult(12)
        Console.WriteLine(t.IsCompleted)
        Console.WriteLine(t.Result)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "12"]);
}

#[test]
fn async_task_matrix_task_run_executes_lambda() {
    let out = run_vb(
        r#"
Imports System.Threading.Tasks

Module M
    Sub Main()
        Dim t As Task(Of Integer) = Task.Run(Function() 2 + 4)
        Console.WriteLine(t.Result)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["6"]);
}

#[test]
fn async_task_matrix_async_function_getawaiter_result() {
    let out = run_vb(
        r#"
Imports System.Threading.Tasks

Module M
    Public Async Function DoubleAsync(ByVal x As Integer) As Task(Of Integer)
        Return Await Task.FromResult(x * 2)
    End Function

    Sub Main()
        Dim t = DoubleAsync(9)
        Console.WriteLine(t.GetAwaiter().GetResult())
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["18"]);
}

#[test]
fn async_task_matrix_when_all_aggregates_results() {
    let out = run_vb(
        r#"
Imports System
Imports System.Linq
Imports System.Threading.Tasks

Module M
    Sub Main()
        Dim a As Task(Of Integer) = Task.Run(Function() 1)
        Dim b As Task(Of Integer) = Task.Run(Function() 2)
        Dim c As Task(Of Integer) = Task.Run(Function() 3)

        Dim all As Task(Of Integer()) = Task.WhenAll(a, b, c)
        Console.WriteLine(all.IsCompleted)
        Console.WriteLine(all.Result.Sum())
        Console.WriteLine(all.Result.Length)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "6", "3"]);
}

#[test]
fn async_task_matrix_when_any_returns_any_completed_task() {
    let out = run_vb(
        r#"
Imports System
Imports System.Threading.Tasks

Module M
    Sub Main()
        Dim fast As Task(Of Integer) = Task.Run(Function() 7)
        Dim slow As Task(Of Integer) = Task.Run(Function()
            Return 9
        End Function)

        Dim winner As Task(Of Integer) = Task.WhenAny(fast, slow).Result
        Console.WriteLine(winner.Status = TaskStatus.RanToCompletion)
        Console.WriteLine(winner.Result = 7 OrElse winner.Result = 9)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn async_task_matrix_continuation_chain() {
    let out = run_vb(
        r#"
Imports System.Threading.Tasks

Module M
    Sub Main()
        Dim baseTask As Task(Of Integer) = Task.Run(Function() 11)
        Dim continued As Task(Of Integer) = baseTask.ContinueWith(Function(t) t.Result + 1)
        Console.WriteLine(continued.Result)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["12"]);
}

#[test]
fn async_task_matrix_factory_start_new_with_string_result() {
    let out = run_vb(
        r#"
Imports System.Threading.Tasks

Module M
    Sub Main()
        Dim t As Task(Of String) = Task.Factory.StartNew(Function() "vb")
        Console.WriteLine(t.Result)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["vb"]);
}

#[test]
fn async_task_matrix_delay_and_wait_guarantees_progress() {
    let out = run_vb(
        r#"
Imports System
Imports System.Threading.Tasks

Module M
    Sub Main()
        Dim t As Task = Task.Delay(1)
        t.Wait()
        Console.WriteLine(t.IsCompleted)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}
