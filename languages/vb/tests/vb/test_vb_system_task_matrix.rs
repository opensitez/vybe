use super::helpers::run_vb;

#[test]
fn task_from_result_returns_value() {
    let out = run_vb(
        r#"
Imports System.Threading.Tasks

Module M
    Sub Main()
        Dim t As Task(Of Integer) = Task.FromResult(9)
        Console.WriteLine(t.IsCompleted)
        Console.WriteLine(t.Result)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "9"]);
}

#[test]
fn task_run_executes_function() {
    let out = run_vb(
        r#"
Imports System.Threading.Tasks

Module M
    Sub Main()
        Dim t As Task(Of Integer) = Task.Run(Function() 2 + 3)
        Console.WriteLine(t.Result)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["5"]);
}

#[test]
fn task_continue_with_applies_transformation() {
    let out = run_vb(
        r#"
Imports System.Threading.Tasks

Module M
    Sub Main()
        Dim seed As Task(Of Integer) = Task.Run(Function() 4)
        Dim mapped As Task(Of Integer) = seed.ContinueWith(Function(t) t.Result * 3)
        Console.WriteLine(mapped.Result)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["12"]);
}

#[test]
fn task_wait_all_collects_results() {
    let out = run_vb(
        r#"
Imports System.Threading.Tasks

Module M
    Sub Main()
        Dim a As Task(Of Integer) = Task.Run(Function() 1)
        Dim b As Task(Of Integer) = Task.Run(Function() 2)
        Dim all As Task(Of Integer()) = Task.WhenAll(a, b)
        Console.WriteLine(all.Result.Length)
        Console.WriteLine(all.Result(0) + all.Result(1))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["2", "3"]);
}

#[test]
fn task_wait_any_returns_completed() {
    let out = run_vb(
        r#"
Imports System.Threading.Tasks

Module M
    Sub Main()
        Dim first As Task(Of Integer) = Task.Run(Function() 7)
        Dim second As Task(Of Integer) = Task.Run(Function() 8)
        Dim winner As Task(Of Integer) = Task.WhenAny(first, second).Result
        Console.WriteLine(winner.Result = 7 OrElse winner.Result = 8)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}

#[test]
fn task_factory_start_new() {
    let out = run_vb(
        r#"
Imports System.Threading.Tasks

Module M
    Sub Main()
        Dim t As Task(Of String) = Task.Factory.StartNew(Function() "ok")
        Console.WriteLine(t.Result)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["ok"]);
}

#[test]
fn task_delay_completes_with_minimum_wait() {
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
