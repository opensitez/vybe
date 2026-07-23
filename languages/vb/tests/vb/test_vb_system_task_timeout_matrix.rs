use super::helpers::run_vb;

#[test]
fn task_timeout_short_wait_can_fail_when_not_ready() {
    let out = run_vb(
        r#"
Imports System.Threading
Imports System.Threading.Tasks

Module M
    Sub Main()
        Dim slow As Task = Task.Run(Sub()
            Thread.Sleep(50)
        End Sub)
        Console.WriteLine(slow.Wait(1))
        Console.WriteLine(slow.Wait(2000))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["False", "True"]);
}

#[test]
fn task_wait_with_infinite_timeout() {
    let out = run_vb(
        r#"
Imports System.Threading
Imports System.Threading.Tasks

Module M
    Sub Main()
        Dim fast As Task(Of Integer) = Task.Run(Function()
            Thread.Sleep(10)
            Return 10
        End Function)
        Console.WriteLine(fast.Wait(Timeout.Infinite))
        Console.WriteLine(fast.Result)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "10"]);
}

#[test]
fn task_waitall_with_timeout() {
    let out = run_vb(
        r#"
Imports System.Threading
Imports System.Threading.Tasks

Module M
    Sub Main()
        Dim a As Task = Task.Run(Sub() Thread.Sleep(10))
        Dim b As Task = Task.Run(Sub() Thread.Sleep(20))
        Dim timeoutExpired As Boolean = Not Task.WaitAll(New Task() {a, b}, 1)
        Dim completed As Boolean = Task.WaitAll(New Task() {a, b}, 2000)
        Console.WriteLine(timeoutExpired)
        Console.WriteLine(completed)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}
