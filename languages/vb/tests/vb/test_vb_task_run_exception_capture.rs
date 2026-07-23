use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Task.Run Exception Unwrapping & AggregateException
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_task_run_uncaught_exception_aggregates() {
    let src = r#"
Imports System
Imports System.Threading.Tasks

Module Program
    Sub Main()
        Dim t = Task.Run(Sub()
            Throw New InvalidOperationException("Task Exception")
        End Sub)
        Try
            t.Wait()
        Catch ex As AggregateException
            Console.WriteLine(ex.InnerException.Message)
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Task Exception"]);
}

#[test]
fn test_vb_task_run_await_unwraps_first_exception() {
    let src = r#"
Imports System
Imports System.Threading.Tasks

Module Program
    Private Async Function RunFaultyTaskAsync() As Task
        Await Task.Run(Sub()
            Throw New ArgumentNullException("param", "Argument null in task")
        End Sub)
    End Function

    Sub Main()
        Try
            Dim t = RunFaultyTaskAsync()
            t.Wait()
        Catch ex As AggregateException
            Dim inner = ex.InnerException
            Console.WriteLine(inner.GetType().Name & ": " & inner.Message)
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["ArgumentNullException: Argument null in task\r\nParameter name: param"]
    );
}

#[test]
fn test_vb_task_run_multiple_exceptions_flatten() {
    let src = r#"
Imports System
Imports System.Threading.Tasks

Module Program
    Sub Main()
        Dim t1 = Task.Run(Sub() Throw New Exception("E1"))
        Dim t2 = Task.Run(Sub() Throw New Exception("E2"))

        Try
            Task.WaitAll(t1, t2)
        Catch ex As AggregateException
            Dim flat = ex.Flatten()
            Console.WriteLine(flat.InnerExceptions.Count)
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2"]);
}

#[test]
fn test_vb_task_run_handle_predicate_selective_catch() {
    let src = r#"
Imports System
Imports System.Threading.Tasks

Module Program
    Sub Main()
        Dim t = Task.Run(Sub() Throw New FormatException("Invalid Format"))
        Try
            t.Wait()
        Catch ex As AggregateException
            ex.Handle(Function(e)
                If TypeOf e Is FormatException Then
                    Console.WriteLine("Handled FormatException: " & e.Message)
                    Return True
                End If
                Return False
            End Function)
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Handled FormatException: Invalid Format"]);
}

#[test]
fn test_vb_task_run_custom_exception_propagation() {
    let src = r#"
Imports System
Imports System.Threading.Tasks

Class CustomDomainException
    Inherits Exception
    Public ErrorCode As Integer
    Public Sub New(code As Integer, msg As String)
        MyBase.New(msg)
        ErrorCode = code
    End Sub
End Class

Module Program
    Sub Main()
        Dim t = Task.Run(Sub()
            Throw New CustomDomainException(404, "Entity Not Found")
        End Sub)
        Try
            t.Wait()
        Catch ex As AggregateException
            Dim cust = CType(ex.InnerException, CustomDomainException)
            Console.WriteLine(cust.ErrorCode & ": " & cust.Message)
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["404: Entity Not Found"]);
}

#[test]
fn test_vb_task_run_is_faulted_property() {
    let src = r#"
Imports System
Imports System.Threading.Tasks

Module Program
    Sub Main()
        Dim t = Task.Run(Sub() Throw New Exception("Error"))
        Try : t.Wait() : Catch : End Try
        Console.WriteLine(t.IsFaulted & "|" & t.IsCompleted)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_task_run_successful_is_not_faulted() {
    let src = r#"
Imports System.Threading.Tasks

Module Program
    Sub Main()
        Dim t = Task.Run(Function() 100)
        t.Wait()
        Console.WriteLine(t.IsFaulted & "|" & (t.Exception Is Nothing))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False|True"]);
}

#[test]
fn test_vb_task_run_operation_canceled_exception_sets_iscanceled() {
    let src = r#"
Imports System
Imports System.Threading.Tasks

Module Program
    Sub Main()
        Dim t = Task.Run(Sub() Throw New OperationCanceledException())
        Try : t.Wait() : Catch : End Try
        Console.WriteLine(t.IsCanceled & "|" & t.IsFaulted)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|False"]);
}

#[test]
fn test_vb_task_continue_with_only_on_faulted() {
    let src = r#"
Imports System
Imports System.Threading.Tasks

Module Program
    Sub Main()
        Dim t = Task.Run(Sub() Throw New Exception("Task Failed"))
        Dim faultTask = t.ContinueWith(
            Sub(ancestor) Console.WriteLine("Fault Handler: " & ancestor.Exception.InnerException.Message),
            TaskContinuationOptions.OnlyOnFaulted
        )
        faultTask.Wait()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Fault Handler: Task Failed"]);
}

#[test]
fn test_vb_task_continue_with_only_on_ran_to_completion() {
    let src = r#"
Imports System.Threading.Tasks

Module Program
    Sub Main()
        Dim t = Task.Run(Function() "SuccessData")
        Dim okTask = t.ContinueWith(
            Sub(ancestor) Console.WriteLine("OK Handler: " & ancestor.Result),
            TaskContinuationOptions.OnlyOnRanToCompletion
        )
        okTask.Wait()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["OK Handler: SuccessData"]);
}

#[test]
fn test_vb_task_run_nested_tasks_exception() {
    let src = r#"
Imports System
Imports System.Threading.Tasks

Module Program
    Sub Main()
        Dim parent = Task.Run(Function()
            Dim child = Task.Run(Sub() Throw New Exception("Nested Exception"))
            child.Wait()
            Return 0
        End Function)

        Try
            parent.Wait()
        Catch ex As AggregateException
            Dim flat = ex.Flatten()
            Console.WriteLine(flat.InnerException.Message)
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Nested Exception"]);
}

#[test]
fn test_vb_task_run_generic_function_exception() {
    let src = r#"
Imports System
Imports System.Threading.Tasks

Module Program
    Private Function ThrowingFunction() As String
        Throw New InvalidCastException("Bad cast in function")
    End Function

    Sub Main()
        Dim t = Task.Run(AddressOf ThrowingFunction)
        Try
            Dim res = t.Result
        Catch ex As AggregateException
            Console.WriteLine(ex.InnerException.GetType().Name)
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["InvalidCastException"]);
}

#[test]
fn test_vb_task_run_divide_by_zero_exception() {
    let src = r#"
Imports System
Imports System.Threading.Tasks

Module Program
    Sub Main()
        Dim t = Task.Run(Function()
            Dim a = 10, b = 0
            Return a \ b
        End Function)

        Try
            Dim r = t.Result
        Catch ex As AggregateException
            Console.WriteLine(ex.InnerException.GetType().Name)
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["DivideByZeroException"]);
}

#[test]
fn test_vb_task_when_all_collects_all_exceptions() {
    let src = r#"
Imports System
Imports System.Threading.Tasks

Module Program
    Private Async Function RunAllAsync() As Task
        Dim t1 = Task.Run(Sub() Throw New InvalidOperationException("E1"))
        Dim t2 = Task.Run(Sub() Throw New ArgumentException("E2"))
        Await Task.WhenAll(t1, t2)
    End Function

    Sub Main()
        Dim mainTask = RunAllAsync()
        Try
            mainTask.Wait()
        Catch ex As AggregateException
            Dim flat = ex.Flatten()
            Console.WriteLine(flat.InnerExceptions.Count)
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2"]);
}

#[test]
fn test_vb_task_run_with_state_object_exception() {
    let src = r#"
Imports System
Imports System.Threading.Tasks

Module Program
    Sub Main()
        Dim t = Task.Factory.StartNew(Sub(state)
            Throw New Exception("State Task Error: " & state.ToString())
        End Sub, "StateValue")

        Try
            t.Wait()
        Catch ex As AggregateException
            Console.WriteLine(ex.InnerException.Message)
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["State Task Error: StateValue"]);
}

#[test]
fn test_vb_task_exception_observed_status() {
    let src = r#"
Imports System
Imports System.Threading.Tasks

Module Program
    Sub Main()
        Dim t = Task.Run(Sub() Throw New Exception("Unobserved Error"))
        Try
            t.Wait()
        Catch
        End Try
        Console.WriteLine(t.IsFaulted & "|" & (t.Exception.InnerException.Message = "Unobserved Error"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_task_run_cancellation_token_parameter() {
    let src = r#"
Imports System.Threading
Imports System.Threading.Tasks

Module Program
    Sub Main()
        Dim cts As New CancellationTokenSource()
        cts.Cancel()
        Dim t = Task.Run(Sub()
            ' Action doesn't run because token was already canceled
        End Sub, cts.Token)

        Console.WriteLine(t.IsCanceled)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_async_lambda_in_task_run_exception() {
    let src = r#"
Imports System
Imports System.Threading.Tasks

Module Program
    Sub Main()
        Dim t = Task.Run(Async Function() As Task
            Await Task.Delay(5)
            Throw New InvalidOperationException("Async Lambda Error")
        End Function)

        Try
            t.Wait()
        Catch ex As AggregateException
            Console.WriteLine(ex.InnerException.Message)
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Async Lambda Error"]);
}

#[test]
fn test_vb_task_run_returning_null_reference_exception() {
    let src = r#"
Imports System
Imports System.Threading.Tasks

Module Program
    Sub Main()
        Dim t = Task.Run(Sub()
            Dim s As String = Nothing
            Dim len = s.Length
        End Sub)

        Try
            t.Wait()
        Catch ex As AggregateException
            Console.WriteLine(ex.InnerException.GetType().Name)
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["NullReferenceException"]);
}

#[test]
fn test_vb_task_exception_property_non_blocking_read() {
    let src = r#"
Imports System
Imports System.Threading.Tasks

Module Program
    Sub Main()
        Dim t = Task.Run(Sub() Throw New Exception("Failure"))
        Try : t.Wait() : Catch : End Try
        Dim aggEx = t.Exception
        Console.WriteLine(aggEx.InnerExceptions(0).Message)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Failure"]);
}
